from __future__ import annotations

import importlib.util
from pathlib import Path
import subprocess
import sys

import pytest


ROOT = Path(__file__).resolve().parents[1]


def _load_runner():
    spec = importlib.util.spec_from_file_location("pptx_wiki_run_script", ROOT / "run.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _prepare_default_project(
    runner, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> tuple[Path, Path, Path, list[str]]:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    (input_dir / "sample.pptx").write_bytes(b"pptx-placeholder")
    config = tmp_path / "config.yml"
    config.write_text("version: 1\n", encoding="utf-8")
    output = tmp_path / "output"
    project_python = tmp_path / ".venv" / "Scripts" / "python.exe"
    recorded: list[str] = []

    monkeypatch.setattr(runner, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(runner, "_project_python", lambda: project_python)

    def fake_run(command) -> int:
        recorded.extend(command)
        return 0

    monkeypatch.setattr(runner, "_run_command", fake_run)
    return input_dir, config, output, recorded


def test_runner_without_arguments_processes_project_input_as_batch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    runner = _load_runner()
    input_dir, config, output, command = _prepare_default_project(
        runner, tmp_path, monkeypatch
    )

    assert runner.main([]) == 0

    assert command[:6] == [
        str(tmp_path / ".venv" / "Scripts" / "python.exe"),
        "-u",
        "-m",
        "pptx_wiki.cli",
        "batch",
        str(input_dir.absolute()),
    ]
    assert command[command.index("--config") + 1] == str(config.absolute())
    assert command[command.index("--output") + 1] == str(output.absolute())
    assert "--recursive" not in command


def test_runner_without_positional_input_honors_recursive_flag(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    runner = _load_runner()
    input_dir, config, output, command = _prepare_default_project(
        runner, tmp_path, monkeypatch
    )

    assert runner.main(["--recursive"]) == 0

    assert command[4:6] == ["batch", str(input_dir.absolute())]
    assert command[command.index("--config") + 1] == str(config.absolute())
    assert command[command.index("--output") + 1] == str(output.absolute())
    assert command.count("--recursive") == 1


def test_runner_resume_quartz_requires_an_explicit_collection(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    runner = _load_runner()
    _prepare_default_project(runner, tmp_path, monkeypatch)

    with pytest.raises(SystemExit, match="2"):
        runner.main(["--resume-quartz"])

    assert "--resume-quartz requires" in capsys.readouterr().err


def test_runner_uses_project_python_and_streams_to_configured_cli(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    runner = _load_runner()
    source = tmp_path / "한글 자료.pptx"
    source.write_bytes(b"pptx-placeholder")
    config = tmp_path / "settings.yml"
    config.write_text("version: 1\n", encoding="utf-8")
    project_python = tmp_path / ".venv" / "Scripts" / "python.exe"
    recorded: dict[str, object] = {}

    monkeypatch.setattr(runner, "_project_python", lambda: project_python)

    def fake_run(command, *, cwd, env, check):
        recorded.update(command=command, cwd=cwd, env=env, check=check)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(runner.subprocess, "run", fake_run)

    assert runner.main([str(source), "--config", str(config)]) == 0
    command = recorded["command"]
    assert command[:5] == [
        str(project_python),
        "-u",
        "-m",
        "pptx_wiki.cli",
        "convert",
    ]
    assert str(source) in command
    assert str(config) in command
    assert recorded["cwd"] == runner.PROJECT_ROOT
    assert recorded["check"] is False
    assert recorded["env"]["PYTHONUNBUFFERED"] == "1"


def test_runner_rejects_non_pptx_input(tmp_path: Path) -> None:
    runner = _load_runner()
    source = tmp_path / "unsafe.pptm"
    source.write_bytes(b"placeholder")
    config = tmp_path / "config.yml"
    config.write_text("version: 1\n", encoding="utf-8")

    with pytest.raises(SystemExit, match="2"):
        runner.main([str(source), "--config", str(config)])


def test_runner_resume_quartz_bypasses_pptx_and_config_validation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    runner = _load_runner()
    collection = tmp_path / "finished-collection"
    collection.mkdir()
    missing_config = tmp_path / "does-not-exist.yml"
    output = tmp_path / "republished-quartz"
    project_python = tmp_path / ".venv" / "Scripts" / "python.exe"
    recorded: dict[str, object] = {}

    monkeypatch.setattr(runner, "_project_python", lambda: project_python)

    def fake_run(command, *, cwd, env, check):
        recorded.update(command=command, cwd=cwd, env=env, check=check)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(runner.subprocess, "run", fake_run)

    assert (
        runner.main(
            [
                str(collection),
                "--resume-quartz",
                "--config",
                str(missing_config),
                "--output",
                str(output),
                "--site-title",
                "Recovered Wiki",
            ]
        )
        == 0
    )

    command = recorded["command"]
    assert command == [
        str(project_python),
        "-u",
        "-m",
        "pptx_wiki.cli",
        "quartz",
        str(collection.absolute()),
        "--site-title",
        "Recovered Wiki",
        "--output",
        str(output.absolute()),
    ]
    assert str(missing_config) not in command
