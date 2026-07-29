"""Download and validate the pinned OvisOCR2 Hugging Face snapshot.

This command is the only part of the worker that needs network access.  The
inference worker never imports this module and always runs with Hugging Face and
Transformers offline modes enabled.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from typing import Any


MODEL_ID = "ATH-MaaS/OvisOCR2"
REVISION = "65c619d374b55d4152e85150fc1b003700bc1f0c"
MODEL_FILENAME = "model.safetensors"
MODEL_SIZE = 1_706_030_496
MODEL_SHA256 = "19e991f6a777a29f6b75d4d02d62c4a5e4ec2f49c94a978f9e134ebc18218b21"
MANIFEST_FILENAME = "pptx-wiki-model-manifest.json"
PROTOCOL_PREFIX = "@@PPTX_WIKI@@"
REQUIRED_FILES = (
    "LICENSE",
    "README.md",
    "chat_template.jinja",
    "config.json",
    "merges.txt",
    MODEL_FILENAME,
    "preprocessor_config.json",
    "tokenizer.json",
    "tokenizer_config.json",
    "video_preprocessor_config.json",
    "vocab.json",
)
ALLOW_PATTERNS = list(REQUIRED_FILES)
TOKEN_ENV_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def default_model_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "models" / "ovisocr2"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(8 * 1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"invalid JSON file {path.name}: {exc}") from exc
    if not isinstance(value, dict):
        raise RuntimeError(f"JSON file must contain an object: {path.name}")
    return value


def validate_snapshot(model_dir: Path) -> dict[str, Any]:
    """Validate the fixed snapshot without executing repository code."""

    missing = [name for name in REQUIRED_FILES if not (model_dir / name).is_file()]
    if missing:
        raise RuntimeError(f"model snapshot is incomplete; missing: {', '.join(missing)}")

    model_path = model_dir / MODEL_FILENAME
    actual_size = model_path.stat().st_size
    if actual_size != MODEL_SIZE:
        raise RuntimeError(
            f"unexpected {MODEL_FILENAME} size: {actual_size}; expected {MODEL_SIZE}"
        )

    config = _read_json(model_dir / "config.json")
    if config.get("model_type") != "qwen3_5":
        raise RuntimeError(f"unexpected model_type: {config.get('model_type')!r}")
    if config.get("architectures") != ["Qwen3_5ForConditionalGeneration"]:
        raise RuntimeError(f"unexpected model architecture: {config.get('architectures')!r}")
    if "auto_map" in config:
        raise RuntimeError("snapshot unexpectedly requests custom remote model code")

    processor = _read_json(model_dir / "preprocessor_config.json")
    if processor.get("processor_class") != "Qwen3VLProcessor":
        raise RuntimeError(f"unexpected processor class: {processor.get('processor_class')!r}")

    # safe_open reads and validates the safetensors header without materializing
    # the 1.7 GB of tensor data in RAM.
    try:
        from safetensors import safe_open

        with safe_open(model_path, framework="pt", device="cpu") as tensors:
            tensor_count = len(list(tensors.keys()))
    except Exception as exc:
        raise RuntimeError(f"invalid safetensors checkpoint: {exc}") from exc
    if tensor_count < 1:
        raise RuntimeError("safetensors checkpoint contains no tensors")

    print("Verifying the pinned model checksum (about 1.7 GB)...", file=sys.stderr)
    model_sha256 = _sha256(model_path)
    if model_sha256 != MODEL_SHA256:
        raise RuntimeError(
            f"unexpected {MODEL_FILENAME} SHA-256: {model_sha256}; expected {MODEL_SHA256}"
        )

    files: dict[str, dict[str, Any]] = {}
    for name in REQUIRED_FILES:
        path = model_dir / name
        files[name] = {
            "size": path.stat().st_size,
            "sha256": model_sha256 if name == MODEL_FILENAME else _sha256(path),
        }
    return {
        "model_type": config["model_type"],
        "architectures": config["architectures"],
        "transformers_version_in_config": config.get("transformers_version"),
        "tensor_count": tensor_count,
        "files": files,
    }


def write_manifest(model_dir: Path, validation: dict[str, Any]) -> Path:
    manifest_path = model_dir / MANIFEST_FILENAME
    manifest = {
        "schema_version": 1,
        "model_id": MODEL_ID,
        "revision": REVISION,
        "license": "Apache-2.0",
        "downloaded_at_utc": datetime.now(timezone.utc).isoformat(),
        "validation": validation,
    }
    temporary = manifest_path.with_suffix(manifest_path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, manifest_path)
    return manifest_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download the pinned ATH-MaaS/OvisOCR2 snapshot for offline inference."
    )
    parser.add_argument(
        "--model-dir",
        type=Path,
        default=default_model_dir(),
        help="Destination directory (default: <repository>/models/ovisocr2).",
    )
    parser.add_argument(
        "--token-env",
        default="HF_TOKEN",
        help="Environment variable containing an optional Hugging Face token. The public model needs none.",
    )
    parser.add_argument(
        "--force-download",
        action="store_true",
        help="Ask huggingface_hub to fetch all files again without deleting the destination.",
    )
    parser.add_argument("--max-workers", type=int, default=4)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not TOKEN_ENV_PATTERN.fullmatch(args.token_env):
        print(f"invalid --token-env name: {args.token_env!r}", file=sys.stderr)
        return 2
    if args.max_workers < 1 or args.max_workers > 16:
        print("--max-workers must be between 1 and 16", file=sys.stderr)
        return 2

    model_dir = args.model_dir.expanduser().resolve()
    if model_dir.exists() and not model_dir.is_dir():
        print(f"model destination is not a directory: {model_dir}", file=sys.stderr)
        return 2
    model_dir.mkdir(parents=True, exist_ok=True)

    os.environ.setdefault("HF_HUB_DISABLE_TELEMETRY", "1")
    os.environ.setdefault("HF_HUB_DISABLE_IMPLICIT_TOKEN", "1")
    token = os.getenv(args.token_env) or None

    try:
        from huggingface_hub import snapshot_download

        downloaded = Path(
            snapshot_download(
                repo_id=MODEL_ID,
                revision=REVISION,
                local_dir=model_dir,
                allow_patterns=ALLOW_PATTERNS,
                token=token,
                force_download=args.force_download,
                max_workers=args.max_workers,
            )
        ).resolve()
        if downloaded != model_dir:
            raise RuntimeError(
                f"huggingface_hub returned an unexpected destination: {downloaded}"
            )
        validation = validate_snapshot(model_dir)
        manifest_path = write_manifest(model_dir, validation)
    except Exception as exc:
        print(f"OvisOCR2 download or validation failed: {exc}", file=sys.stderr)
        print(
            PROTOCOL_PREFIX
            + json.dumps(
                {
                    "ok": False,
                    "model_id": MODEL_ID,
                    "revision": REVISION,
                    "model_dir": str(model_dir),
                    "error": str(exc),
                },
                ensure_ascii=False,
            )
        )
        return 1

    print(
        PROTOCOL_PREFIX
        + json.dumps(
            {
                "ok": True,
                "model_id": MODEL_ID,
                "revision": REVISION,
                "model_dir": str(model_dir),
                "manifest": str(manifest_path),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
