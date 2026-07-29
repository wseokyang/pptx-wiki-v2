from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys

import pytest


ROOT = Path(__file__).resolve().parents[1]
PROFILES = (
    "paddleocr_vl_16",
    "paddleocr_vl_15",
    "monkeyocr_v2_b",
    "ovisocr2",
)
PROTOCOL = "pptx-wiki-ocr-worker/1"
PREFIX = "@@PPTX_WIKI@@"


@pytest.mark.parametrize("profile", PROFILES)
def test_worker_reports_machine_readable_startup_failure_without_model(
    profile: str, tmp_path: Path
) -> None:
    model_dir = tmp_path / profile
    model_dir.mkdir()
    completed = subprocess.run(
        [
            sys.executable,
            str(ROOT / "workers" / profile / "worker.py"),
            "--serve",
            "--model-dir",
            str(model_dir),
            "--device",
            "auto",
            "--dtype",
            "auto",
            "--max-new-tokens",
            "16384",
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=10,
        check=False,
    )

    assert completed.returncode != 0
    frames = [
        json.loads(line.split(PREFIX, 1)[1])
        for line in completed.stdout.splitlines()
        if PREFIX in line
    ]
    assert len(frames) == 1, (completed.stdout, completed.stderr)
    assert frames[0]["protocol"] == PROTOCOL
    assert frames[0]["type"] == "ready"
    assert frames[0]["ok"] is False
    assert frames[0]["error"]["message"]


@pytest.mark.parametrize("profile", PROFILES)
def test_windows_setup_downloads_by_default_and_can_skip(profile: str) -> None:
    setup = (ROOT / "workers" / profile / "setup-windows.ps1").read_text(
        encoding="utf-8"
    )

    assert "[switch]$SkipDownload" in setup
    assert "if (-not $SkipDownload)" in setup
    assert "download.py" in setup


@pytest.mark.parametrize("profile", PROFILES)
def test_profile_has_an_independent_locked_environment(profile: str) -> None:
    directory = ROOT / "workers" / profile

    for name in (
        "README.md",
        "download.py",
        "requirements.lock.txt",
        "setup-windows.ps1",
        "worker.py",
    ):
        assert (directory / name).is_file(), f"{profile} is missing {name}"
    assert ".venv" in (directory / "setup-windows.ps1").read_text(encoding="utf-8")
