"""Create the local application environment for this repository.

Packaging metadata intentionally stays in ``pyproject.toml``.  This file is a
repository bootstrapper, not a legacy setuptools ``setup.py`` entry point.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Sequence
import venv


PROJECT_ROOT = Path(__file__).resolve().parent
MINIMUM_PYTHON = (3, 10)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create or reuse .venv, install pptx-wiki, and initialize the local "
            "config/input/output workspace."
        )
    )
    parser.add_argument(
        "--dev",
        action="store_true",
        help="also install development/test dependencies",
    )
    parser.add_argument(
        "--skip-install",
        action="store_true",
        help="only create/validate .venv and initialize config.yml",
    )
    parser.add_argument(
        "--skip-config",
        action="store_true",
        help="do not create config.yml from config.example.yml",
    )
    return parser


def _environment_dir() -> Path:
    return PROJECT_ROOT / ".venv"


def _venv_python(root: Path) -> Path:
    if sys.platform == "win32":
        return root / "Scripts" / "python.exe"
    return root / "bin" / "python"


def _require_supported_python() -> None:
    if sys.version_info[:2] < MINIMUM_PYTHON:
        required = ".".join(str(value) for value in MINIMUM_PYTHON)
        actual = ".".join(str(value) for value in sys.version_info[:3])
        raise RuntimeError(f"Python {required} or newer is required (found {actual})")


def _ensure_environment() -> Path:
    """Create the project venv once, or validate and reuse it unchanged."""

    root = _environment_dir()
    python = _venv_python(root)
    if root.exists():
        if (
            not root.is_dir()
            or not (root / "pyvenv.cfg").is_file()
            or not python.is_file()
        ):
            raise RuntimeError(
                f"{root} exists but is not a complete virtual environment; "
                "move or remove it after checking its contents"
            )
        return python

    venv.EnvBuilder(with_pip=True).create(root)
    if not (root / "pyvenv.cfg").is_file() or not python.is_file():
        raise RuntimeError(f"virtual environment creation did not complete: {root}")
    return python


def _run_checked(command: Sequence[str]) -> None:
    subprocess.run(
        list(command),
        cwd=PROJECT_ROOT,
        check=True,
        shell=False,
    )


def _install(python: Path, *, dev: bool) -> None:
    extras = "api,windows,dev" if dev else "api,windows"
    _run_checked(
        [str(python), "-m", "pip", "install", "-e", f".[{extras}]"]
    )
    _run_checked([str(python), "-m", "pip", "check"])
    _run_checked(
        [
            str(python),
            "-c",
            "import pptx_wiki; print('pptx-wiki', pptx_wiki.__version__)",
        ]
    )


def _initialize_config() -> bool:
    """Create config.yml exclusively; never replace a user's configuration."""

    destination = PROJECT_ROOT / "config.yml"
    if destination.exists():
        if not destination.is_file():
            raise RuntimeError(f"configuration path is not a regular file: {destination}")
        return False
    source = PROJECT_ROOT / "config.example.yml"
    if not source.is_file():
        raise FileNotFoundError(f"configuration template was not found: {source}")
    try:
        with source.open("rb") as input_handle, destination.open("xb") as output_handle:
            shutil.copyfileobj(input_handle, output_handle)
    except FileExistsError:
        return False
    return True


def _initialize_io_directories() -> None:
    """Create the local input/output directories without replacing user data."""

    directories = (PROJECT_ROOT / "input", PROJECT_ROOT / "output")
    for directory in directories:
        if directory.exists() and not directory.is_dir():
            raise RuntimeError(
                f"workspace path exists but is not a directory: {directory}"
            )
    for directory in directories:
        directory.mkdir(exist_ok=True)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        _require_supported_python()
        python = _ensure_environment()
        print(f"Project environment ready: {_environment_dir()}", flush=True)
        if not args.skip_install:
            _install(python, dev=bool(args.dev))
        config_created = False
        if not args.skip_config:
            config_created = _initialize_config()
            print(
                f"{'Created' if config_created else 'Kept existing'} configuration: "
                f"{PROJECT_ROOT / 'config.yml'}",
                flush=True,
            )
        _initialize_io_directories()
        print(f"Input directory ready : {PROJECT_ROOT / 'input'}", flush=True)
        print(f"Output directory ready: {PROJECT_ROOT / 'output'}", flush=True)
    except subprocess.CalledProcessError as error:
        print(
            f"bootstrap failed: command exited with status {error.returncode}",
            file=sys.stderr,
        )
        return error.returncode or 1
    except (OSError, RuntimeError) as error:
        print(f"bootstrap failed: {error}", file=sys.stderr)
        return 1

    print("", flush=True)
    print("Setup complete.", flush=True)
    print("1. Edit config.yml and set llm_api endpoint/model credentials.", flush=True)
    print(
        "2. Put .pptx files in input/.",
        flush=True,
    )
    print(
        "3. Run: python run.py",
        flush=True,
    )
    print(
        "4. PR/request numbers must be present in native text or tables while "
        "image extraction is disabled.",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
