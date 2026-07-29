from __future__ import annotations

import argparse
import importlib.util
import os
from pathlib import Path
import subprocess
import sys
from typing import Sequence


PROJECT_ROOT = Path(__file__).resolve().parent


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert one PPTX into a grounded LLM wiki using config.yml.",
    )
    parser.add_argument("input", type=Path, help="input .pptx file")
    parser.add_argument(
        "--config",
        type=Path,
        default=PROJECT_ROOT / "config.yml",
        help="YAML configuration (default: config.yml next to this script)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="override the output directory configured in YAML",
    )
    return parser


def _absolute(path: Path) -> Path:
    return path.expanduser().resolve()


def _project_python() -> Path:
    candidates = (
        PROJECT_ROOT / ".venv" / "Scripts" / "python.exe",
        PROJECT_ROOT / ".venv" / "bin" / "python",
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    if importlib.util.find_spec("pptx_wiki") is not None:
        return Path(sys.executable).resolve()
    raise FileNotFoundError(
        "project Python environment was not found; run setup-windows.ps1 once "
        "or select an interpreter where pptx-wiki is installed"
    )


def _command(input_path: Path, config_path: Path, output_path: Path | None) -> list[str]:
    command = [
        str(_project_python()),
        "-u",
        "-m",
        "pptx_wiki.cli",
        "convert",
        str(input_path),
        "--config",
        str(config_path),
    ]
    if output_path is not None:
        command.extend(("--output", str(output_path)))
    return command


def main(argv: Sequence[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    input_path = _absolute(args.input)
    config_path = _absolute(args.config)
    output_path = _absolute(args.output) if args.output is not None else None

    if not input_path.is_file():
        parser.error(f"PPTX file not found: {input_path}")
    if input_path.suffix.casefold() != ".pptx":
        parser.error(f"only macro-free .pptx files are accepted: {input_path}")
    if not config_path.is_file():
        parser.error(
            f"configuration file not found: {config_path}; copy config.example.yml "
            "to config.yml and edit the endpoint settings"
        )

    try:
        command = _command(input_path, config_path, output_path)
    except FileNotFoundError as exc:
        parser.error(str(exc))

    print(f"PPTX   : {input_path}", flush=True)
    print(f"Config : {config_path}", flush=True)
    print(f"Python : {command[0]}", flush=True)
    print("-" * 72, flush=True)

    environment = os.environ.copy()
    environment.setdefault("PYTHONUTF8", "1")
    environment["PYTHONUNBUFFERED"] = "1"
    try:
        completed = subprocess.run(
            command,
            cwd=PROJECT_ROOT,
            env=environment,
            check=False,
        )
    except KeyboardInterrupt:
        return 130
    except OSError as exc:
        print(f"pptx-wiki launcher error: {exc}", file=sys.stderr)
        return 2
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
