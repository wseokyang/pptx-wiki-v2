"""Download and verify the pinned MonkeyOCRv2-B-Parsing snapshot."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from model_spec import MANIFEST_FILENAME, MODEL_REPO_ID, MODEL_REVISION, REQUIRED_FILES


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(8 * 1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _check_required(
    model_dir: Path,
    verified: dict[str, tuple[int, str]] | None = None,
) -> dict[str, str]:
    errors: dict[str, str] = {}
    for relative, (expected_size, expected_hash) in REQUIRED_FILES.items():
        path = model_dir / relative
        cached = verified.get(relative) if verified is not None else None
        if cached is not None:
            actual_size, actual_hash = cached
        else:
            if not path.is_file():
                errors[relative] = "missing"
                continue
            actual_size = path.stat().st_size
            actual_hash = ""
        if actual_size != expected_size:
            errors[relative] = f"size {actual_size}, expected {expected_size}"
            continue
        if not actual_hash:
            actual_hash = sha256_file(path)
        if actual_hash.lower() != expected_hash.lower():
            errors[relative] = f"sha256 {actual_hash}, expected {expected_hash}"
    return errors


def _snapshot_files(model_dir: Path) -> list[Path]:
    files: list[Path] = []
    for path in model_dir.rglob("*"):
        if not path.is_file() or path.name == MANIFEST_FILENAME:
            continue
        relative = path.relative_to(model_dir)
        if relative.parts and relative.parts[0] == ".cache":
            continue
        files.append(path)
    return sorted(files, key=lambda path: path.relative_to(model_dir).as_posix())


def build_manifest(model_dir: Path) -> dict[str, Any]:
    records = []
    for path in _snapshot_files(model_dir):
        relative = path.relative_to(model_dir).as_posix()
        records.append(
            {
                "path": relative,
                "size": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
    return {
        "schema_version": 1,
        "repo_id": MODEL_REPO_ID,
        "revision": MODEL_REVISION,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "total_bytes": sum(record["size"] for record in records),
        "required_files": sorted(REQUIRED_FILES),
        "files": records,
    }


def write_manifest(model_dir: Path, manifest: dict[str, Any]) -> Path:
    destination = model_dir / MANIFEST_FILENAME
    fd, temporary_name = tempfile.mkstemp(
        prefix=f".{MANIFEST_FILENAME}.", suffix=".tmp", dir=model_dir
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(manifest, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary_name, destination)
    except Exception:
        try:
            os.unlink(temporary_name)
        except OSError:
            pass
        raise
    return destination


def verify_manifest(model_dir: Path) -> list[str]:
    manifest_path = model_dir / MANIFEST_FILENAME
    if not manifest_path.is_file():
        return [f"missing {MANIFEST_FILENAME}"]
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return [f"invalid manifest: {exc}"]
    errors: list[str] = []
    if manifest.get("repo_id") != MODEL_REPO_ID:
        errors.append("manifest repo_id does not match the pinned model")
    if manifest.get("revision") != MODEL_REVISION:
        errors.append("manifest revision does not match the pinned revision")
    records = manifest.get("files")
    if not isinstance(records, list):
        return errors + ["manifest files must be an array"]
    root = model_dir.resolve()
    verified: dict[str, tuple[int, str]] = {}
    for record in records:
        if not isinstance(record, dict) or not isinstance(record.get("path"), str):
            errors.append("manifest contains an invalid file record")
            continue
        relative = Path(record["path"])
        path = (model_dir / relative).resolve()
        try:
            path.relative_to(root)
        except ValueError:
            errors.append(f"manifest path escapes model directory: {relative}")
            continue
        if not path.is_file():
            errors.append(f"missing manifest file: {relative.as_posix()}")
            continue
        expected_size = record.get("size")
        if not isinstance(expected_size, int) or path.stat().st_size != expected_size:
            errors.append(f"size mismatch: {relative.as_posix()}")
            continue
        expected_hash = record.get("sha256")
        actual_hash = sha256_file(path)
        if not isinstance(expected_hash, str) or actual_hash != expected_hash.lower():
            errors.append(f"sha256 mismatch: {relative.as_posix()}")
            continue
        verified[relative.as_posix()] = (path.stat().st_size, actual_hash)
    errors.extend(
        f"{name}: {reason}" for name, reason in _check_required(model_dir, verified).items()
    )
    return errors


def parse_args() -> argparse.Namespace:
    worker_root = Path(__file__).resolve().parent
    repo_root = worker_root.parent.parent
    parser = argparse.ArgumentParser(
        description="Download the exact, pinned MonkeyOCRv2-B-Parsing snapshot."
    )
    parser.add_argument(
        "--model-dir",
        type=Path,
        default=repo_root / "models" / "monkeyocr_v2_b",
    )
    parser.add_argument("--token-env", default="HF_TOKEN")
    parser.add_argument("--force", action="store_true", help="Force Hugging Face to redownload files.")
    parser.add_argument("--verify-only", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    model_dir = args.model_dir.expanduser().resolve()
    if args.verify_only:
        errors = verify_manifest(model_dir)
        if errors:
            print("Verification failed:", file=sys.stderr)
            for error in errors:
                print(f"  - {error}", file=sys.stderr)
            return 2
        print(f"Verified pinned model at {model_dir}")
        return 0

    model_dir.mkdir(parents=True, exist_ok=True)
    free_bytes = shutil.disk_usage(model_dir).free
    if free_bytes < 3 * 1024**3:
        print(
            "Warning: less than 3 GiB is free; the full snapshot and temporary files may not fit.",
            file=sys.stderr,
        )

    try:
        from huggingface_hub import snapshot_download
    except ImportError as exc:
        print("huggingface-hub is missing; run setup-windows.ps1 first.", file=sys.stderr)
        return 2

    os.environ.setdefault("HF_HUB_DISABLE_TELEMETRY", "1")
    token = os.environ.get(args.token_env) or None
    print(f"Downloading {MODEL_REPO_ID}@{MODEL_REVISION} to {model_dir}")
    snapshot_download(
        repo_id=MODEL_REPO_ID,
        revision=MODEL_REVISION,
        local_dir=model_dir,
        token=token,
        force_download=args.force,
    )

    required_errors = _check_required(model_dir)
    if required_errors:
        print("Pinned-file verification failed:", file=sys.stderr)
        for name, reason in required_errors.items():
            print(f"  - {name}: {reason}", file=sys.stderr)
        print("Run again with --force after checking the download source.", file=sys.stderr)
        return 2

    manifest_path = write_manifest(model_dir, build_manifest(model_dir))
    errors = verify_manifest(model_dir)
    if errors:
        print("Final manifest verification failed:", file=sys.stderr)
        for error in errors:
            print(f"  - {error}", file=sys.stderr)
        return 2
    print(f"Verified model and wrote {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
