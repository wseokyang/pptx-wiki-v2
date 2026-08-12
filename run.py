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
        description=(
            "Process one PPTX or a collection of PPTX files using config.yml."
        ),
    )
    parser.add_argument(
        "input",
        nargs="+",
        type=Path,
        help="one or more input .pptx files, or directories in --batch mode",
    )
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
    parser.add_argument(
        "--batch",
        action="store_true",
        help="run parsed -> semantic -> integrated -> Quartz for all inputs",
    )
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--site-title", default="신뢰성 분석 LLM Wiki")
    return parser


def _absolute(path: Path) -> Path:
    # Preserve lexical components so the collection preflight can reject a
    # symlink/junction instead of silently resolving through it.
    return path.expanduser().absolute()


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
        "project Python environment was not found; run 'python bootstrap.py' once "
        "or select an interpreter where pptx-wiki is installed"
    )


def _command(
    input_paths: Sequence[Path],
    config_path: Path,
    output_path: Path | None,
    *,
    batch: bool,
    recursive: bool = False,
    site_title: str = "신뢰성 분석 LLM Wiki",
) -> list[str]:
    command = [
        str(_project_python()),
        "-u",
        "-m",
        "pptx_wiki.cli",
        "batch" if batch else "convert",
        *(str(path) for path in input_paths),
        "--config",
        str(config_path),
    ]
    if output_path is not None:
        command.extend(("--output", str(output_path)))
    if batch:
        command.extend(("--site-title", site_title))
        if recursive:
            command.append("--recursive")
    return command


def main(argv: Sequence[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    input_paths = tuple(_absolute(value) for value in args.input)
    config_path = _absolute(args.config)
    output_path = _absolute(args.output) if args.output is not None else None
    batch = bool(args.batch or len(input_paths) > 1 or any(path.is_dir() for path in input_paths))

    for input_path in input_paths:
        if not input_path.exists():
            parser.error(f"input not found: {input_path}")
        if input_path.is_file() and input_path.suffix.casefold() != ".pptx":
            parser.error(f"only macro-free .pptx files are accepted: {input_path}")
        if not input_path.is_file() and not (batch and input_path.is_dir()):
            parser.error(f"input must be a .pptx file or batch directory: {input_path}")
    if batch and output_path is None:
        parser.error("--output is required in batch mode")
    if not config_path.is_file():
        parser.error(
            f"configuration file not found: {config_path}; copy config.example.yml "
            "to config.yml and edit the endpoint settings"
        )

    try:
        command = _command(
            input_paths,
            config_path,
            output_path,
            batch=batch,
            recursive=bool(args.recursive),
            site_title=args.site_title,
        )
    except FileNotFoundError as exc:
        parser.error(str(exc))

    print(f"Input  : {', '.join(str(path) for path in input_paths)}", flush=True)
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
