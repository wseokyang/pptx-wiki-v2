"""Offline PaddleOCR-VL 1.6 worker with one-shot and persistent JSONL modes."""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Mapping


# These must be set before importing PaddleOCR/PaddleX. All model directories
# are supplied explicitly, so inference must never consult a remote model host.
os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["HF_DATASETS_OFFLINE"] = "1"
os.environ["HF_HUB_DISABLE_TELEMETRY"] = "1"
os.environ["HF_HUB_DISABLE_IMPLICIT_TOKEN"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"
os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

FRAME_PREFIX = "@@PPTX_WIKI@@"
PROTOCOL = "pptx-wiki-ocr-worker/1"
WORKER_NAME = "paddleocr_vl_16"
PIPELINE_VERSION = "v1.6"
VL_MODEL_NAME = "PaddleOCR-VL-1.6-0.9B"
WORKER_ROOT = Path(__file__).resolve().parent
DEFAULT_MODEL_DIR = WORKER_ROOT.parent.parent / "models" / WORKER_NAME
ALLOWED_TASKS = frozenset({"document", "text", "table", "chart", "formula"})
PROMPT_LABELS = {
    "text": "ocr",
    "table": "table",
    "chart": "chart",
    "formula": "formula",
}
EXPECTED_MODELS = {
    "vlm": {
        "repo_id": "PaddlePaddle/PaddleOCR-VL-1.6",
        "revision": "66317acc4c9fc17bd154591ce650735cd2855f3e",
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
    },
    "layout": {
        "repo_id": "PaddlePaddle/PP-DocLayoutV3",
        "revision": "7b48a7566925fa464281f930c58eee04fe2c862a",
        "directory": "layout",
        "required_files": (
            "inference.json",
            "inference.pdiparams",
            "inference.yml",
        ),
    },
}
FORBIDDEN_MODEL_SUFFIXES = frozenset(
    {".bin", ".pkl", ".pickle", ".pt", ".pth", ".py"}
)


def _normalize_device(device: str | None) -> str | None:
    """Translate the shared worker device syntax to Paddle's device syntax."""
    if device is None:
        return None
    normalized = device.strip().casefold()
    if not normalized or normalized == "auto":
        # Omitting the option lets Paddle prefer GPU 0 and fall back to CPU.
        return None
    if normalized in {"cuda", "gpu"}:
        return "gpu:0"
    if normalized.startswith("cuda:"):
        index = normalized.removeprefix("cuda:")
        if not index.isdecimal():
            raise ValueError("CUDA device must be 'cuda' or 'cuda:<non-negative index>'")
        return f"gpu:{int(index)}"
    return normalized


def _emit(payload: Mapping[str, Any]) -> None:
    framed = dict(payload)
    framed["protocol"] = PROTOCOL
    line = json.dumps(framed, ensure_ascii=False, separators=(",", ":"))
    print(f"{FRAME_PREFIX}{line}", flush=True)


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
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


def _validate_models(models_dir: Path) -> tuple[Path, Path]:
    models_dir = models_dir.resolve()
    manifest_path = models_dir / "manifest.json"
    if not manifest_path.is_file():
        raise RuntimeError(
            f"model manifest is missing: {manifest_path}; run download.py first"
        )
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"invalid model manifest: {manifest_path}") from exc
    if manifest.get("worker") != WORKER_NAME:
        raise RuntimeError(f"model manifest belongs to another worker: {manifest_path}")

    entries = {
        item.get("role"): item
        for item in manifest.get("models", [])
        if isinstance(item, Mapping)
    }
    resolved: dict[str, Path] = {}
    for role, expected in EXPECTED_MODELS.items():
        entry = entries.get(role)
        if not isinstance(entry, Mapping):
            raise RuntimeError(f"manifest has no {role!r} model entry")
        if entry.get("repo_id") != expected["repo_id"]:
            raise RuntimeError(f"unexpected {role} repository in model manifest")
        if entry.get("revision") != expected["revision"]:
            raise RuntimeError(f"unexpected {role} revision in model manifest")
        model_dir = models_dir / str(expected["directory"])
        for name in expected["required_files"]:
            path = model_dir / str(name)
            if not path.is_file() or path.stat().st_size <= 0:
                raise RuntimeError(f"required model file is missing or empty: {path}")
        forbidden = sorted(
            path
            for path in model_dir.rglob("*")
            if path.is_file() and path.suffix.casefold() in FORBIDDEN_MODEL_SUFFIXES
        )
        if forbidden:
            raise RuntimeError(
                f"forbidden executable or pickle-style model file found: {forbidden[0]}"
            )
        resolved[role] = model_dir
    return resolved["vlm"], resolved["layout"]


class PaddleWorker:
    def __init__(
        self,
        *,
        model_dir: Path,
        device: str | None,
        dtype: str | None,
        max_new_tokens: int | None,
    ) -> None:
        vlm_dir, layout_dir = _validate_models(model_dir)
        try:
            from paddleocr import PaddleOCRVL
        except ImportError as exc:
            raise RuntimeError(
                "PaddleOCR is not installed; run setup-windows.ps1 first"
            ) from exc

        normalized_device = _normalize_device(device)
        options: dict[str, Any] = {
            "pipeline_version": PIPELINE_VERSION,
            "layout_detection_model_name": "PP-DocLayoutV3",
            "layout_detection_model_dir": str(layout_dir),
            "vl_rec_model_name": VL_MODEL_NAME,
            "vl_rec_model_dir": str(vlm_dir),
            "vl_rec_backend": "native",
            "use_doc_orientation_classify": False,
            "use_doc_unwarping": False,
            "use_layout_detection": True,
            "format_block_content": True,
            "merge_layout_blocks": False,
            "use_queues": False,
        }
        if normalized_device is not None:
            options["device"] = normalized_device
        self.pipeline = PaddleOCRVL(**options)
        self.device = normalized_device
        # PaddleOCRVL chooses its native dtype from the selected device. The
        # argument is accepted for a stable cross-worker CLI contract.
        self.dtype = dtype
        self.max_new_tokens = max_new_tokens

    def recognize(self, request: Mapping[str, Any]) -> dict[str, Any]:
        request_id = request.get("id")
        if request_id is not None and not isinstance(request_id, (str, int)):
            raise ValueError("id must be a string, integer, or null")
        image_value = request.get("image")
        image_path_value = request.get("image_path")
        if (
            image_value is not None
            and image_path_value is not None
            and image_value != image_path_value
        ):
            raise ValueError("image and image_path must match when both are provided")
        if image_value is None:
            image_value = image_path_value
        if not isinstance(image_value, str) or not image_value.strip():
            raise ValueError(
                "image or image_path must be a non-empty local path string"
            )
        image_path = Path(image_value).expanduser().resolve()
        if not image_path.is_file():
            raise ValueError(f"image file does not exist: {image_path}")
        task = request.get("task", "document")
        if not isinstance(task, str) or task not in ALLOWED_TASKS:
            raise ValueError(f"task must be one of {sorted(ALLOWED_TASKS)}")
        language = request.get("language", "ko")
        if not isinstance(language, str):
            raise ValueError("language must be a string")
        context = request.get("context")
        if context is not None and not isinstance(context, str):
            raise ValueError("context must be a string or null")

        predict_options: dict[str, Any] = {
            "format_block_content": True,
            "use_queues": False,
        }
        if self.max_new_tokens is not None:
            predict_options["max_new_tokens"] = self.max_new_tokens
        if task == "document":
            predict_options.update(
                {
                    "use_layout_detection": True,
                    "merge_layout_blocks": False,
                    "layout_merge_bboxes_mode": "union",
                }
            )
        else:
            predict_options.update(
                {
                    "use_layout_detection": False,
                    "prompt_label": PROMPT_LABELS[task],
                }
            )

        pages = list(self.pipeline.predict(str(image_path), **predict_options))
        if len(pages) != 1:
            raise RuntimeError(
                f"one image must produce exactly one result, received {len(pages)}"
            )
        page = pages[0]
        page_json = page.json
        if not isinstance(page_json, Mapping) or not isinstance(
            page_json.get("res"), Mapping
        ):
            raise RuntimeError("PaddleOCR returned an unexpected JSON result")
        result = dict(page_json["res"])
        try:
            markdown = page.markdown
            if isinstance(markdown, Mapping):
                markdown_text = markdown.get("markdown_texts")
                if isinstance(markdown_text, str):
                    result["markdown"] = markdown_text
        except Exception as exc:
            result.setdefault("warnings", []).append(
                f"Markdown conversion failed: {type(exc).__name__}: {exc}"
            )
        result["worker_metadata"] = {
            "worker": WORKER_NAME,
            "pipeline_version": PIPELINE_VERSION,
            "vl_model_name": VL_MODEL_NAME,
            "task": task,
            "language_hint": language,
            "device": self.device or "auto",
            "requested_dtype": self.dtype,
            "max_new_tokens": self.max_new_tokens,
        }
        return result


def _result_envelope(
    worker: PaddleWorker, request: Mapping[str, Any]
) -> dict[str, Any]:
    request_id = request.get("id")
    try:
        result = worker.recognize(request)
    except Exception as exc:
        traceback.print_exc(file=sys.stderr)
        return {
            "type": "result",
            "id": request_id,
            "ok": False,
            "error": {"type": type(exc).__name__, "message": str(exc)},
        }
    return {"type": "result", "id": request_id, "ok": True, "result": result}


def _serve_jsonl(worker: PaddleWorker) -> int:
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
            if not isinstance(request, Mapping):
                raise ValueError("each JSONL request must be an object")
            message_type = request.get("type")
            if message_type in {"request", "shutdown", "ping"}:
                operation = {
                    "request": "recognize",
                    "shutdown": "shutdown",
                    "ping": "ping",
                }[message_type]
            else:
                operation = request.get("op", "recognize")
            if operation == "shutdown":
                _emit(
                    {
                        "type": "result",
                        "id": request.get("id"),
                        "ok": True,
                        "shutdown": True,
                    }
                )
                return 0
            if operation == "ping":
                _emit(
                    {
                        "type": "result",
                        "id": request.get("id"),
                        "ok": True,
                        "pong": True,
                    }
                )
                continue
            if operation != "recognize":
                raise ValueError(f"unsupported operation: {operation!r}")
            response = _result_envelope(worker, request)
        except Exception as exc:
            traceback.print_exc(file=sys.stderr)
            response = {
                "type": "result",
                "id": None,
                "ok": False,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            }
        _emit(response)
    return 0


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="backslashreplace")


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument(
        "--serve",
        "--jsonl",
        dest="serve",
        action="store_true",
        help="serve newline-delimited JSON on stdin",
    )
    mode.add_argument("--image", type=str, help="run one local image request")
    parser.add_argument(
        "--model-dir",
        "--models-dir",
        dest="model_dir",
        type=Path,
        default=DEFAULT_MODEL_DIR,
    )
    parser.add_argument(
        "--device",
        help="device: auto, cpu, gpu:0, cuda, or cuda:0 (cuda aliases map to Paddle gpu)",
    )
    parser.add_argument(
        "--dtype",
        help="accepted for worker compatibility; native Paddle selects the actual dtype",
    )
    parser.add_argument("--max-new-tokens", type=int)
    parser.add_argument("--task", choices=sorted(ALLOWED_TASKS), default="document")
    parser.add_argument("--id", default="one-shot")
    parser.add_argument("--language", default="ko")
    parser.add_argument("--context")
    parser.add_argument(
        "--output",
        type=Path,
        help="also write the one-shot result envelope to JSON",
    )
    args = parser.parse_args(argv)

    try:
        if args.max_new_tokens is not None and args.max_new_tokens < 1:
            parser.error("--max-new-tokens must be at least 1")
        worker = PaddleWorker(
            model_dir=args.model_dir,
            device=args.device,
            dtype=args.dtype,
            max_new_tokens=args.max_new_tokens,
        )
    except Exception as exc:
        traceback.print_exc(file=sys.stderr)
        _emit(
            {
                "type": "ready",
                "ok": False,
                "worker": WORKER_NAME,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            }
        )
        return 1

    _emit(
        {
            "type": "ready",
            "ok": True,
            "worker": WORKER_NAME,
            "pipeline_version": PIPELINE_VERSION,
            "device": worker.device or "auto",
            "dtype": args.dtype,
            "max_new_tokens": args.max_new_tokens,
        }
    )
    if args.serve:
        return _serve_jsonl(worker)

    request = {
        "id": args.id,
        "image": args.image,
        "task": args.task,
        "language": args.language,
        "context": args.context,
    }
    response = _result_envelope(worker, request)
    _emit(response)
    if args.output is not None:
        _atomic_write_json(args.output.resolve(), response)
    return 0 if response.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
