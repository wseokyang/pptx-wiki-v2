"""Download the exact PaddleOCR-VL 1.5 pipeline models from Hugging Face."""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any


WORKER_NAME = "paddleocr_vl_15"
WORKER_ROOT = Path(__file__).resolve().parent
DEFAULT_MODEL_DIR = WORKER_ROOT.parent.parent / "models" / WORKER_NAME

MODEL_SPECS = (
    {
        "role": "vlm",
        "repo_id": "PaddlePaddle/PaddleOCR-VL-1.5",
        "revision": "426bf5b6c89670e370e71ce0c51cf2bb458b7db9",
        "directory": "vlm",
        "required_files": (
            "added_tokens.json",
            "chat_template.jinja",
            "config.json",
            "generation_config.json",
            "inference.yml",
            "model.safetensors",
            "preprocessor_config.json",
            "processor_config.json",
            "special_tokens_map.json",
            "tokenizer.json",
            "tokenizer.model",
            "tokenizer_config.json",
        ),
        "allow_patterns": (
            "LICENSE",
            "added_tokens.json",
            "chat_template.jinja",
            "config.json",
            "generation_config.json",
            "inference.yml",
            "model.safetensors",
            "preprocessor_config.json",
            "processor_config.json",
            "special_tokens_map.json",
            "tokenizer.json",
            "tokenizer.model",
            "tokenizer_config.json",
        ),
    },
    {
        "role": "layout",
        "repo_id": "PaddlePaddle/PP-DocLayoutV3",
        "revision": "7b48a7566925fa464281f930c58eee04fe2c862a",
        "directory": "layout",
        "required_files": (
            "inference.json",
            "inference.pdiparams",
            "inference.yml",
        ),
        "allow_patterns": (
            "LICENSE",
            "inference.json",
            "inference.pdiparams",
            "inference.yml",
        ),
    },
)

FORBIDDEN_MODEL_SUFFIXES = frozenset(
    {".bin", ".pkl", ".pickle", ".pt", ".pth", ".py"}
)


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        "w",
        encoding="utf-8",
        newline="\n",
        dir=path.parent,
        prefix=f".{path.name}.",
        suffix=".tmp",
        delete=False,
    ) as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
        temp_path = Path(handle.name)
    os.replace(temp_path, path)


def _validate_files(model_dir: Path, required_files: tuple[str, ...]) -> dict[str, int]:
    sizes: dict[str, int] = {}
    for relative_name in required_files:
        path = model_dir / relative_name
        if not path.is_file():
            raise RuntimeError(f"required model file is missing: {path}")
        size = path.stat().st_size
        if size <= 0:
            raise RuntimeError(f"required model file is empty: {path}")
        sizes[relative_name] = size
    return sizes


def _validate_no_executable_model_files(model_dir: Path) -> None:
    forbidden = sorted(
        path
        for path in model_dir.rglob("*")
        if path.is_file() and path.suffix.casefold() in FORBIDDEN_MODEL_SUFFIXES
    )
    if forbidden:
        joined = ", ".join(str(path) for path in forbidden[:5])
        raise RuntimeError(
            "model directory contains executable or pickle-style files; remove the "
            f"directory and download it again: {joined}"
        )


def download(model_dir: Path, *, force: bool = False) -> Path:
    try:
        from huggingface_hub import snapshot_download
    except ImportError as exc:
        raise RuntimeError(
            "huggingface-hub is not installed; run setup-windows.ps1 first"
        ) from exc

    model_dir = model_dir.resolve()
    model_dir.mkdir(parents=True, exist_ok=True)
    token = os.environ.get("HF_TOKEN") or None
    entries: list[dict[str, Any]] = []

    for spec in MODEL_SPECS:
        target = model_dir / str(spec["directory"])
        target.mkdir(parents=True, exist_ok=True)
        print(
            f"Downloading {spec['repo_id']}@{spec['revision']} to {target}",
            flush=True,
        )
        snapshot_download(
            repo_id=str(spec["repo_id"]),
            revision=str(spec["revision"]),
            local_dir=target,
            token=token,
            force_download=force,
            allow_patterns=list(spec["allow_patterns"]),
            ignore_patterns=["*.bin", "*.pkl", "*.pickle", "*.pt", "*.pth", "*.py"],
        )
        required = tuple(str(item) for item in spec["required_files"])
        sizes = _validate_files(target, required)
        _validate_no_executable_model_files(target)
        entries.append(
            {
                "role": spec["role"],
                "repo_id": spec["repo_id"],
                "revision": spec["revision"],
                "directory": spec["directory"],
                "required_files": sizes,
                "allow_patterns": list(spec["allow_patterns"]),
            }
        )

    manifest = {
        "schema_version": 1,
        "worker": WORKER_NAME,
        "downloaded_at": datetime.now(timezone.utc).isoformat(),
        "models": entries,
    }
    manifest_path = model_dir / "manifest.json"
    _atomic_write_json(manifest_path, manifest)
    return manifest_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--model-dir",
        "--models-dir",
        dest="model_dir",
        type=Path,
        default=DEFAULT_MODEL_DIR,
        help="destination directory (default: repository models/paddleocr_vl_15)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="ask Hugging Face Hub to redownload files even when cached",
    )
    args = parser.parse_args(argv)
    try:
        manifest_path = download(args.model_dir, force=args.force)
    except Exception as exc:
        print(f"download failed: {exc}", file=sys.stderr)
        return 1
    print(f"Model manifest written to {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
