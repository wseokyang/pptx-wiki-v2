from __future__ import annotations

import importlib.util
from pathlib import Path
import subprocess
import sys

import pytest


ROOT = Path(__file__).resolve().parents[1]


def _load_bootstrap():
    spec = importlib.util.spec_from_file_location(
        "pptx_wiki_bootstrap_script", ROOT / "bootstrap.py"
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_bootstrap_main_runs_default_install_and_initializes_config(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    environment_python = tmp_path / ".venv" / "Scripts" / "python.exe"
    calls: list[tuple[object, ...]] = []

    def fake_environment() -> Path:
        calls.append(("environment",))
        return environment_python

    monkeypatch.setattr(bootstrap, "_ensure_environment", fake_environment)
    monkeypatch.setattr(
        bootstrap,
        "_install",
        lambda python, *, dev=False: calls.append(("install", python, dev)),
    )
    monkeypatch.setattr(
        bootstrap, "_initialize_config", lambda: calls.append(("config",))
    )
    monkeypatch.setattr(
        bootstrap,
        "_initialize_io_directories",
        lambda: calls.append(("io-directories",)),
    )

    assert bootstrap.main([]) == 0
    assert calls == [
        ("environment",),
        ("install", environment_python, False),
        ("config",),
        ("io-directories",),
    ]


def test_bootstrap_main_honors_dev_and_skip_flags(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    environment_python = tmp_path / ".venv" / "Scripts" / "python.exe"
    calls: list[tuple[object, ...]] = []

    def fake_environment() -> Path:
        calls.append(("environment",))
        return environment_python

    monkeypatch.setattr(bootstrap, "_ensure_environment", fake_environment)
    monkeypatch.setattr(
        bootstrap,
        "_install",
        lambda python, *, dev=False: calls.append(("install", python, dev)),
    )
    monkeypatch.setattr(
        bootstrap, "_initialize_config", lambda: calls.append(("config",))
    )
    monkeypatch.setattr(
        bootstrap,
        "_initialize_io_directories",
        lambda: calls.append(("io-directories",)),
    )

    assert bootstrap.main(["--dev", "--skip-config"]) == 0
    assert calls == [
        ("environment",),
        ("install", environment_python, True),
        ("io-directories",),
    ]

    calls.clear()
    assert bootstrap.main(["--skip-install", "--skip-config"]) == 0
    assert calls == [("environment",), ("io-directories",)]


def test_bootstrap_creates_project_environment(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)
    environment = tmp_path / ".venv"
    created: list[Path] = []

    class FakeBuilder:
        def __init__(self, **kwargs) -> None:
            assert kwargs["with_pip"] is True

        def create(self, path: Path) -> None:
            created.append(Path(path))
            (Path(path) / "pyvenv.cfg").parent.mkdir(parents=True, exist_ok=True)
            (Path(path) / "pyvenv.cfg").write_text(
                "home = test-python\n", encoding="utf-8"
            )
            python = bootstrap._venv_python(Path(path))
            python.parent.mkdir(parents=True, exist_ok=True)
            python.touch()

    monkeypatch.setattr(bootstrap.venv, "EnvBuilder", FakeBuilder)

    expected_python = bootstrap._venv_python(environment)
    assert bootstrap._ensure_environment() == expected_python
    assert created == [environment]


def test_bootstrap_reuses_valid_project_environment_without_recreating_it(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)
    environment = tmp_path / ".venv"
    (environment / "pyvenv.cfg").parent.mkdir(parents=True)
    (environment / "pyvenv.cfg").write_text(
        "home = test-python\n", encoding="utf-8"
    )
    expected_python = bootstrap._venv_python(environment)
    expected_python.parent.mkdir(parents=True, exist_ok=True)
    expected_python.touch()

    class UnexpectedBuilder:
        def __init__(self, **kwargs) -> None:
            pytest.fail(f"a valid existing venv must be reused: {kwargs}")

    monkeypatch.setattr(bootstrap.venv, "EnvBuilder", UnexpectedBuilder)

    assert bootstrap._ensure_environment() == expected_python


def test_bootstrap_rejects_broken_project_environment_without_modifying_it(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)
    environment = tmp_path / ".venv"
    environment.mkdir()
    marker = environment / "keep-me.txt"
    marker.write_text("user data", encoding="utf-8")

    class UnexpectedBuilder:
        def __init__(self, **kwargs) -> None:
            pytest.fail(f"a broken existing venv must not be overwritten: {kwargs}")

    monkeypatch.setattr(bootstrap.venv, "EnvBuilder", UnexpectedBuilder)

    with pytest.raises(RuntimeError, match="(?i)(invalid|broken|incomplete|python)"):
        bootstrap._ensure_environment()

    assert marker.read_text(encoding="utf-8") == "user data"


@pytest.mark.parametrize(
    ("dev", "requirement"),
    ((False, ".[api,windows]"), (True, ".[api,windows,dev]")),
)
def test_bootstrap_installs_expected_editable_extras(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    dev: bool,
    requirement: str,
) -> None:
    bootstrap = _load_bootstrap()
    python = tmp_path / ".venv" / "python"
    recorded: list[tuple[list[str], dict[str, object]]] = []

    def fake_run(command, **kwargs):
        recorded.append((command, kwargs))
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(bootstrap.subprocess, "run", fake_run)
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)

    bootstrap._install(python, dev=dev)

    assert len(recorded) == 3
    install_command, options = recorded[0]
    assert install_command == [
        str(python),
        "-m",
        "pip",
        "install",
        "-e",
        requirement,
    ]
    assert options["cwd"] == tmp_path
    assert options["check"] is True
    assert options["shell"] is False
    assert all(
        call_options["cwd"] == tmp_path
        and call_options["check"] is True
        and call_options["shell"] is False
        for _, call_options in recorded
    )


def test_bootstrap_initializes_config_once_without_overwriting_user_changes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    template = tmp_path / "config.example.yml"
    config = tmp_path / "config.yml"
    template.write_text("version: 1\nmodel: example\n", encoding="utf-8")
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)

    bootstrap._initialize_config()
    assert config.read_text(encoding="utf-8") == template.read_text(encoding="utf-8")

    config.write_text("version: 1\nmodel: private\n", encoding="utf-8")
    template.write_text("version: 2\nmodel: changed\n", encoding="utf-8")

    bootstrap._initialize_config()
    assert config.read_text(encoding="utf-8") == "version: 1\nmodel: private\n"


def test_bootstrap_initializes_input_and_output_directories_idempotently(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bootstrap = _load_bootstrap()
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)

    bootstrap._initialize_io_directories()

    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    assert input_dir.is_dir()
    assert output_dir.is_dir()

    input_marker = input_dir / "existing.pptx"
    output_marker = output_dir / "existing" / "result.md"
    input_marker.write_bytes(b"existing input")
    output_marker.parent.mkdir()
    output_marker.write_text("existing output\n", encoding="utf-8")

    bootstrap._initialize_io_directories()

    assert input_marker.read_bytes() == b"existing input"
    assert output_marker.read_text(encoding="utf-8") == "existing output\n"


@pytest.mark.parametrize("conflicting_name", ("input", "output"))
def test_bootstrap_rejects_workspace_directory_name_collisions_without_changes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    conflicting_name: str,
) -> None:
    bootstrap = _load_bootstrap()
    monkeypatch.setattr(bootstrap, "PROJECT_ROOT", tmp_path)
    conflict = tmp_path / conflicting_name
    conflict.write_text("user data\n", encoding="utf-8")

    with pytest.raises(RuntimeError, match="not a directory"):
        bootstrap._initialize_io_directories()

    assert conflict.read_text(encoding="utf-8") == "user data\n"
    other_name = "output" if conflicting_name == "input" else "input"
    assert not (tmp_path / other_name).exists()


def test_readme_documents_python_first_bootstrap_and_execution_flow() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")

    bootstrap_position = readme.index("python bootstrap.py")
    run_positions = [
        position
        for command in ("python run.py", "python .\\run.py", "python ./run.py")
        if (position := readme.find(command)) >= 0
    ]

    assert run_positions, "README must show how to run the pipeline with Python"
    assert bootstrap_position < min(run_positions)
    assert "python bootstrap.py --dev" in readme
    assert (
        "python run.py input --batch --recursive --config config.yml --output output"
        in readme
    )
