"""Offline native-Transformers worker for the pinned OvisOCR2 checkpoint.

The persistent ``serve`` command loads the model once and exchanges one JSON
object per line over stdin/stdout.  stdout is reserved for protocol objects;
all diagnostics and third-party output are redirected to stderr.
"""

from __future__ import annotations

import argparse
import contextlib
from dataclasses import dataclass
import hashlib
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
import sys
import traceback
from typing import Any, TextIO


PROTOCOL = "pptx-wiki-ocr-worker/1"
PROTOCOL_PREFIX = "@@PPTX_WIKI@@"
BACKEND = "hf_ovisocr2"
MODEL_ID = "ATH-MaaS/OvisOCR2"
REVISION = "65c619d374b55d4152e85150fc1b003700bc1f0c"
MODEL_FILENAME = "model.safetensors"
MODEL_SIZE = 1_706_030_496
MODEL_SHA256 = "19e991f6a777a29f6b75d4d02d62c4a5e4ec2f49c94a978f9e134ebc18218b21"
MANIFEST_FILENAME = "pptx-wiki-model-manifest.json"
REQUIRED_FILES = (
    "LICENSE",
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
TASKS = frozenset({"document", "text", "table", "chart", "formula"})
LANGUAGE_PATTERN = re.compile(r"^[A-Za-z]{2,8}(?:[-_][A-Za-z0-9]{2,8})?$")
MAX_REQUEST_LINE_CHARS = 128 * 1024
MAX_CONTEXT_CHARS = 8 * 1024
MAX_IMAGE_PIXELS = 100_000_000

OFFICIAL_PROMPT = (
    "Extract all readable content from the image in natural human reading order and output the result "
    "as a single Markdown document. For charts or images, represent them using an HTML image tag: "
    '<img src="images/bbox_{left}_{top}_{right}_{bottom}.jpg" />, where left, top, right, bottom are '
    "bounding box coordinates scaled to [0, 1000). Format formulas as LaTeX. Format tables as HTML: "
    "<table>...</table>. Transcribe all other text as standard Markdown. Preserve the original text "
    "without translation or paraphrasing."
)

TASK_SUFFIXES = {
    "document": "",
    "text": (
        " The input is an isolated text region. Transcribe every readable character verbatim and do not "
        "summarize or correct spelling."
    ),
    "table": (
        " The input is an isolated crop containing exactly one table. Keep it separate from all other "
        "tables and preserve empty cells, row spans, column spans, numbers, signs, percentages, and units."
    ),
    "chart": (
        " The input is an isolated chart region. Preserve every readable title, legend, axis label, category, "
        "and displayed value; do not infer values that are not visibly printed."
    ),
    "formula": (
        " The input is an isolated formula region. Transcribe the formula as LaTeX and preserve adjacent "
        "visible labels verbatim."
    ),
}

TABLE_OR_IMAGE_PATTERN = re.compile(
    r"(?P<table><table\b[^>]*>.*?</table>)"
    r"|(?P<image><img\s+src=[\"']images/bbox_"
    r"(?P<left>\d+)_(?P<top>\d+)_(?P<right>\d+)_(?P<bottom>\d+)\.jpg[\"']\s*/?>)",
    flags=re.IGNORECASE | re.DOTALL,
)


def _set_offline_environment() -> None:
    values = {
        "HF_HUB_OFFLINE": "1",
        "TRANSFORMERS_OFFLINE": "1",
        "HF_DATASETS_OFFLINE": "1",
        "HF_HUB_DISABLE_TELEMETRY": "1",
        "HF_HUB_DISABLE_IMPLICIT_TOKEN": "1",
        "HF_HUB_DISABLE_PROGRESS_BARS": "1",
        "TOKENIZERS_PARALLELISM": "false",
    }
    for name, value in values.items():
        os.environ[name] = value


_set_offline_environment()

# Preserve the real stdout before redirecting third-party imports/model loading.
_PROTOCOL_STREAM: TextIO = sys.stdout


def _emit(value: dict[str, Any]) -> None:
    payload = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
    _PROTOCOL_STREAM.write(PROTOCOL_PREFIX + payload + "\n")
    _PROTOCOL_STREAM.flush()


def _log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


def default_model_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "models" / "ovisocr2"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(8 * 1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_object(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"invalid JSON file {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise RuntimeError(f"JSON file must contain an object: {path}")
    return value


def validate_model_dir(path: Path) -> Path:
    model_dir = path.expanduser().resolve()
    if not model_dir.is_dir():
        raise RuntimeError(f"model directory does not exist: {model_dir}")
    missing = [name for name in REQUIRED_FILES if not (model_dir / name).is_file()]
    if missing:
        raise RuntimeError(f"model snapshot is incomplete; missing: {', '.join(missing)}")

    manifest_path = model_dir / MANIFEST_FILENAME
    if not manifest_path.is_file():
        raise RuntimeError(
            f"missing {MANIFEST_FILENAME}; install this snapshot with download.py before inference"
        )
    manifest = _load_object(manifest_path)
    if manifest.get("schema_version") != 1:
        raise RuntimeError("unsupported or missing model manifest schema")
    if manifest.get("model_id") != MODEL_ID or manifest.get("revision") != REVISION:
        raise RuntimeError("model manifest does not match the pinned OvisOCR2 snapshot")
    validation = manifest.get("validation")
    manifest_files = validation.get("files") if isinstance(validation, dict) else None
    if not isinstance(manifest_files, dict):
        raise RuntimeError("model manifest has no validated file inventory")

    config = _load_object(model_dir / "config.json")
    if config.get("model_type") != "qwen3_5":
        raise RuntimeError(f"unexpected model_type: {config.get('model_type')!r}")
    if config.get("architectures") != ["Qwen3_5ForConditionalGeneration"]:
        raise RuntimeError(f"unexpected architecture: {config.get('architectures')!r}")
    if "auto_map" in config:
        raise RuntimeError("snapshot unexpectedly requests custom remote code")

    checkpoint = model_dir / MODEL_FILENAME
    if checkpoint.stat().st_size != MODEL_SIZE:
        raise RuntimeError(f"unexpected checkpoint size: {checkpoint.stat().st_size}")
    _log("Verifying the pinned OvisOCR2 checkpoint checksum...")
    actual_hash = _sha256(checkpoint)
    if actual_hash != MODEL_SHA256:
        raise RuntimeError(
            f"checkpoint SHA-256 mismatch: {actual_hash}; expected {MODEL_SHA256}"
        )
    for name in REQUIRED_FILES:
        record = manifest_files.get(name)
        if not isinstance(record, dict):
            raise RuntimeError(f"model manifest has no record for {name}")
        file_path = model_dir / name
        expected_size = record.get("size")
        expected_hash = record.get("sha256")
        if file_path.stat().st_size != expected_size:
            raise RuntimeError(f"model file changed after validation: {name} (size mismatch)")
        file_hash = actual_hash if name == MODEL_FILENAME else _sha256(file_path)
        if file_hash != expected_hash:
            raise RuntimeError(f"model file changed after validation: {name} (SHA-256 mismatch)")
    return model_dir


class _HTMLTextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in {"br", "p", "tr", "li", "div", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")
        elif tag.lower() in {"td", "th"}:
            self.parts.append("\t")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"p", "tr", "li", "div", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if data:
            self.parts.append(data)

    def text(self) -> str:
        value = "".join(self.parts)
        value = re.sub(r"[ \t]+\n", "\n", value)
        value = re.sub(r"\n[ \t]+", "\n", value)
        value = re.sub(r"\n{3,}", "\n\n", value)
        return value.strip()


def _plain_text(markup: str) -> str:
    value = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", markup)
    value = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"^\s{0,3}(?:#{1,6}|[-*+]\s|\d+[.)]\s)\s*", "", value, flags=re.MULTILINE)
    value = re.sub(r"(?<!\\)[*_~`]", "", value)
    parser = _HTMLTextExtractor()
    try:
        parser.feed(value)
        parser.close()
        return parser.text()
    except Exception:
        return re.sub(r"<[^>]+>", " ", value).strip()


def _clean_truncated_repeats(
    text: str,
    min_text_len: int = 8000,
    max_period: int = 200,
    min_period: int = 1,
    min_repeat_chars: int = 100,
    min_repeat_times: int = 5,
) -> str:
    """Official OvisOCR2 model-card cleanup for a repeating truncated tail."""

    n = len(text)
    if n < min_text_len:
        return text
    max_period = min(max_period, n - 1)
    for unit_len in range(min_period, max_period + 1):
        if text[n - 1] != text[n - 1 - unit_len]:
            continue
        match_len = 1
        index = n - 2
        while index >= unit_len and text[index] == text[index - unit_len]:
            match_len += 1
            index -= 1
        total_len = match_len + unit_len
        repeat_times = total_len // unit_len
        tail_len = total_len % unit_len
        if repeat_times >= min_repeat_times and total_len >= min_repeat_chars:
            return text[: n - total_len + unit_len] + text[n - tail_len :]
    return text


def _blocks_from_markdown(markdown: str, task: str, width: int, height: int) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    cursor = 0

    def append_text(value: str) -> None:
        value = value.strip()
        if not value:
            return
        kind = task if task in {"text", "chart", "formula"} else "text"
        blocks.append(
            {
                "kind": kind,
                "text": _plain_text(value),
                "markdown": value,
                "html": None,
                "bbox": None,
                "confidence": None,
                "order": len(blocks),
            }
        )

    for match in TABLE_OR_IMAGE_PATTERN.finditer(markdown):
        append_text(markdown[cursor : match.start()])
        if match.group("table") is not None:
            table = match.group("table").strip()
            blocks.append(
                {
                    "kind": "table",
                    "text": _plain_text(table),
                    "markdown": table,
                    "html": table,
                    "bbox": None,
                    "confidence": None,
                    "order": len(blocks),
                }
            )
        else:
            left = int(match.group("left"))
            top = int(match.group("top"))
            right = int(match.group("right"))
            bottom = int(match.group("bottom"))
            bbox = [
                max(0.0, min(float(width), left * width / 1000.0)),
                max(0.0, min(float(height), top * height / 1000.0)),
                max(0.0, min(float(width), right * width / 1000.0)),
                max(0.0, min(float(height), bottom * height / 1000.0)),
            ]
            blocks.append(
                {
                    "kind": "image",
                    "text": "",
                    "markdown": match.group("image"),
                    "html": match.group("image"),
                    "bbox": bbox,
                    "confidence": None,
                    "order": len(blocks),
                    "coordinate_source": "ovis_normalized_0_1000_converted_to_roi_pixels",
                }
            )
        cursor = match.end()
    append_text(markdown[cursor:])

    if not blocks and markdown.strip():
        kind = task if task != "document" else "text"
        blocks.append(
            {
                "kind": kind,
                "text": _plain_text(markdown),
                "markdown": markdown,
                "html": markdown if task == "table" and "<table" in markdown.lower() else None,
                "bbox": None,
                "confidence": None,
                "order": 0,
            }
        )
    return blocks


@dataclass(slots=True)
class Runtime:
    torch: Any
    Image: Any
    processor: Any
    model: Any
    model_dir: Path
    device: Any
    dtype_name: str
    max_new_tokens: int
    min_pixels: int
    max_pixels: int

    def recognize(self, request: dict[str, Any]) -> dict[str, Any]:
        image_path = Path(request["image"]).expanduser().resolve()
        if not image_path.is_file():
            raise RequestError("image_not_found", f"image does not exist: {image_path}")

        task = request["task"]
        language = request["language"]
        prompt = OFFICIAL_PROMPT + TASK_SUFFIXES[task]
        if language:
            language_label = {"ko": "Korean (ko)", "en": "English (en)", "zh": "Chinese (zh)"}.get(
                language.lower(), language
            )
            prompt += (
                f" Expected language hint: {language_label}. Keep that language and do not translate it."
            )

        try:
            with self.Image.open(image_path) as opened:
                opened.load()
                width, height = opened.size
                if width < 1 or height < 1:
                    raise RequestError("invalid_image", "image has empty dimensions")
                if width * height > MAX_IMAGE_PIXELS:
                    raise RequestError(
                        "image_too_large",
                        f"image contains {width * height} pixels; limit is {MAX_IMAGE_PIXELS}",
                    )
                image = opened.convert("RGB")
        except RequestError:
            raise
        except Exception as exc:
            raise RequestError("invalid_image", f"unable to decode image: {exc}") from exc

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "image", "image": image},
                    {"type": "text", "text": prompt},
                ],
            }
        ]
        try:
            inputs = self.processor.apply_chat_template(
                messages,
                add_generation_prompt=True,
                enable_thinking=False,
                tokenize=True,
                return_dict=True,
                return_tensors="pt",
                processor_kwargs={
                    "images_kwargs": {
                        "min_pixels": self.min_pixels,
                        "max_pixels": self.max_pixels,
                    }
                },
            ).to(self.device)
            input_length = int(inputs["input_ids"].shape[-1])
            with self.torch.inference_mode():
                generated = self.model.generate(
                    **inputs,
                    max_new_tokens=self.max_new_tokens,
                    do_sample=False,
                    use_cache=True,
                )
            generated_only = generated[:, input_length:]
            markdown = self.processor.batch_decode(
                generated_only,
                skip_special_tokens=True,
                clean_up_tokenization_spaces=False,
            )[0].strip()
            markdown = _clean_truncated_repeats(markdown)
            hit_token_limit = int(generated_only.shape[-1]) >= self.max_new_tokens
        except self.torch.cuda.OutOfMemoryError as exc:
            self.torch.cuda.empty_cache()
            raise RequestError(
                "cuda_out_of_memory",
                "CUDA ran out of memory; lower max_pixels/max_new_tokens or use a GPU with more VRAM",
            ) from exc
        except RequestError:
            raise
        except Exception as exc:
            raise RequestError("inference_failed", f"OvisOCR2 inference failed: {exc}") from exc

        blocks = _blocks_from_markdown(markdown, task, width, height)
        table_html = [block["html"] for block in blocks if block["kind"] == "table" and block["html"]]
        warnings: list[str] = []
        if hit_token_limit:
            warnings.append(
                f"generation reached max_new_tokens={self.max_new_tokens}; the OCR result may be truncated"
            )
        if not markdown:
            warnings.append("OvisOCR2 returned an empty result")
        return {
            "text": _plain_text(markdown),
            "markdown": markdown,
            "html": "\n".join(table_html) if table_html else None,
            "confidence": None,
            "blocks": blocks,
            "warnings": warnings,
            "metadata": {
                "model_id": MODEL_ID,
                "revision": REVISION,
                "image_width": width,
                "image_height": height,
                "language_hint": language,
                "context_received": bool(request.get("context")),
                "context_used_as_prompt": False,
            },
        }


class RequestError(RuntimeError):
    def __init__(self, code: str, message: str, *, retryable: bool = False) -> None:
        super().__init__(message)
        self.code = code
        self.retryable = retryable


def _resolve_device(torch: Any, requested: str) -> Any:
    requested = requested.strip().lower()
    if requested == "gpu":
        requested = "cuda"
    elif requested.startswith("gpu:"):
        requested = "cuda:" + requested.split(":", 1)[1]
    if requested == "auto":
        return torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
    if requested == "cpu":
        return torch.device("cpu")
    if not re.fullmatch(r"cuda(?::\d+)?", requested):
        raise RuntimeError("--device must be auto, cpu, cuda[:index], or gpu[:index]")
    if not torch.cuda.is_available():
        raise RuntimeError("CUDA was requested but torch.cuda.is_available() is false")
    device = torch.device(requested)
    index = device.index if device.index is not None else 0
    if index >= torch.cuda.device_count():
        raise RuntimeError(f"CUDA device index {index} is unavailable")
    torch.cuda.set_device(index)
    return torch.device(f"cuda:{index}")


def _resolve_dtype(torch: Any, device: Any, requested: str) -> Any:
    mapping = {
        "float32": torch.float32,
        "float16": torch.float16,
        "bfloat16": torch.bfloat16,
    }
    if requested != "auto":
        if device.type == "cpu" and requested == "float16":
            raise RuntimeError("float16 is not supported for this worker on CPU")
        return mapping[requested]
    if device.type == "cuda":
        return torch.bfloat16 if torch.cuda.is_bf16_supported() else torch.float16
    return torch.float32


def load_runtime(args: argparse.Namespace) -> Runtime:
    model_dir = validate_model_dir(args.model_dir)
    _log("Loading OvisOCR2 with local_files_only=True; network access is disabled.")
    with contextlib.redirect_stdout(sys.stderr):
        import torch
        from PIL import Image
        import transformers
        from transformers import AutoModelForMultimodalLM, AutoProcessor

        transformers.utils.logging.set_verbosity_error()
        transformers.utils.logging.disable_progress_bar()
        device = _resolve_device(torch, args.device)
        dtype = _resolve_dtype(torch, device, args.dtype)
        if device.type == "cuda":
            torch.set_float32_matmul_precision("high")
            torch.backends.cuda.matmul.allow_tf32 = True

        processor = AutoProcessor.from_pretrained(
            model_dir,
            local_files_only=True,
            trust_remote_code=False,
        )
        model = AutoModelForMultimodalLM.from_pretrained(
            model_dir,
            local_files_only=True,
            trust_remote_code=False,
            dtype=dtype,
            attn_implementation="sdpa",
            use_kernels=False,
        )
        model.to(device)
        model.eval()

    return Runtime(
        torch=torch,
        Image=Image,
        processor=processor,
        model=model,
        model_dir=model_dir,
        device=device,
        dtype_name=str(dtype).removeprefix("torch."),
        max_new_tokens=args.max_new_tokens,
        min_pixels=args.min_pixels,
        max_pixels=args.max_pixels,
    )


def _validate_request(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise RequestError("invalid_request", "request must be a JSON object")
    request_id = value.get("id")
    if not isinstance(request_id, (str, int)) or isinstance(request_id, bool):
        raise RequestError("invalid_request", "request id must be a string or integer")
    if isinstance(request_id, str) and (not request_id or len(request_id) > 512):
        raise RequestError("invalid_request", "string request id must contain 1 to 512 characters")
    supplied_protocol = value.get("protocol")
    if supplied_protocol is not None and supplied_protocol != PROTOCOL:
        raise RequestError("invalid_request", f"unsupported protocol: {supplied_protocol!r}")
    image = value.get("image") or value.get("image_path")
    if value.get("image") and value.get("image_path") and value["image"] != value["image_path"]:
        raise RequestError("invalid_request", "image and image_path refer to different values")
    if not isinstance(image, str) or not image or len(image) > 32_768:
        raise RequestError("invalid_request", "image must be a non-empty filesystem path")
    task = value.get("task", "document")
    if task not in TASKS:
        raise RequestError("invalid_request", f"unsupported task: {task!r}")
    language = value.get("language", "ko")
    if not isinstance(language, str) or not LANGUAGE_PATTERN.fullmatch(language.strip()):
        raise RequestError(
            "invalid_request",
            "language must be a short BCP-47-like code such as ko, en, or ko-KR",
        )
    context = value.get("context")
    if context is not None and (not isinstance(context, str) or len(context) > MAX_CONTEXT_CHARS):
        raise RequestError(
            "invalid_request", f"context must be null or a string up to {MAX_CONTEXT_CHARS} characters"
        )
    return {
        "id": request_id,
        "image": image,
        "task": task,
        "language": language.strip(),
        "context": context,
    }


def _result_message(runtime: Runtime, raw_request: Any) -> dict[str, Any]:
    request_id = raw_request.get("id") if isinstance(raw_request, dict) else None
    try:
        request = _validate_request(raw_request)
        request_id = request["id"]
        # Preserve stdout exclusively for protocol frames even if an inference
        # dependency unexpectedly calls print().
        with contextlib.redirect_stdout(sys.stderr):
            result = runtime.recognize(request)
        return {
            "protocol": PROTOCOL,
            "type": "result",
            "id": request_id,
            "ok": True,
            "result": result,
        }
    except RequestError as exc:
        _log(f"Request {request_id!r} failed [{exc.code}]: {exc}")
        return {
            "protocol": PROTOCOL,
            "type": "result",
            "id": request_id,
            "ok": False,
            "error": {
                "type": "RequestError",
                "code": exc.code,
                "message": str(exc),
                "retryable": exc.retryable,
            },
        }
    except Exception as exc:  # defensive protocol boundary
        _log(f"Unexpected request failure for {request_id!r}: {exc}")
        traceback.print_exc(file=sys.stderr)
        return {
            "protocol": PROTOCOL,
            "type": "result",
            "id": request_id,
            "ok": False,
            "error": {
                "type": type(exc).__name__,
                "code": "internal_error",
                "message": "unexpected OvisOCR2 worker failure",
                "retryable": False,
            },
        }


def _ready_message(runtime: Runtime) -> dict[str, Any]:
    cuda_name = None
    if runtime.device.type == "cuda":
        cuda_name = runtime.torch.cuda.get_device_name(runtime.device.index or 0)
    return {
        "protocol": PROTOCOL,
        "type": "ready",
        "ok": True,
        "worker": "ovisocr2",
        "backend": BACKEND,
        "model_id": MODEL_ID,
        "revision": REVISION,
        "model_dir": str(runtime.model_dir),
        "device": str(runtime.device),
        "cuda_device": cuda_name,
        "dtype": runtime.dtype_name,
        "min_pixels": runtime.min_pixels,
        "max_pixels": runtime.max_pixels,
        "max_new_tokens": runtime.max_new_tokens,
        "pid": os.getpid(),
    }


def serve(runtime: Runtime) -> int:
    _emit(_ready_message(runtime))
    for line in sys.stdin:
        if not line.strip():
            continue
        if len(line) > MAX_REQUEST_LINE_CHARS:
            _emit(
                {
                    "protocol": PROTOCOL,
                    "type": "result",
                    "id": None,
                    "ok": False,
                    "error": {
                        "type": "RequestError",
                        "code": "request_too_large",
                        "message": f"request line exceeds {MAX_REQUEST_LINE_CHARS} characters",
                        "retryable": False,
                    },
                }
            )
            continue
        try:
            raw = json.loads(line)
        except json.JSONDecodeError as exc:
            _emit(
                {
                    "protocol": PROTOCOL,
                    "type": "result",
                    "id": None,
                    "ok": False,
                    "error": {
                        "type": "JSONDecodeError",
                        "code": "invalid_json",
                        "message": f"invalid request JSON: {exc}",
                        "retryable": False,
                    },
                }
            )
            continue
        if isinstance(raw, dict):
            message_type = raw.get("type")
            operation = {
                "request": raw.get("op", "recognize"),
                "shutdown": "shutdown",
                "ping": "ping",
            }.get(message_type, raw.get("op", "recognize"))
            if operation == "shutdown":
                _emit(
                    {
                        "protocol": PROTOCOL,
                        "type": "result",
                        "id": raw.get("id"),
                        "ok": True,
                        "shutdown": True,
                    }
                )
                _log("Shutdown request received.")
                return 0
            if operation == "ping":
                _emit(
                    {
                        "protocol": PROTOCOL,
                        "type": "result",
                        "id": raw.get("id"),
                        "ok": True,
                        "pong": True,
                    }
                )
                continue
            if operation != "recognize":
                _emit(
                    {
                        "protocol": PROTOCOL,
                        "type": "result",
                        "id": raw.get("id"),
                        "ok": False,
                        "error": {
                            "type": "RequestError",
                            "code": "unsupported_operation",
                            "message": f"unsupported operation: {operation!r}",
                            "retryable": False,
                        },
                    }
                )
                continue
        _emit(_result_message(runtime, raw))
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run pinned OvisOCR2 with native Transformers offline.")
    parser.add_argument(
        "--serve",
        action="store_true",
        help="Load once and read persistent JSONL requests from stdin.",
    )
    parser.add_argument("--model-dir", type=Path, default=default_model_dir())
    parser.add_argument(
        "--device",
        default="auto",
        help="auto, cpu, cuda[:index], or gpu[:index] (gpu is normalized to cuda)",
    )
    parser.add_argument(
        "--dtype",
        choices=("auto", "float32", "float16", "bfloat16"),
        default="auto",
    )
    parser.add_argument("--min-pixels", type=int, default=448 * 448)
    parser.add_argument("--max-pixels", type=int, default=2880 * 2880)
    parser.add_argument("--max-new-tokens", type=int, default=16_384)
    parser.add_argument("--id", default="one-shot")
    parser.add_argument("--image", type=Path, default=None, help="One-shot image; omit with --serve.")
    parser.add_argument("--task", choices=sorted(TASKS), default="document")
    parser.add_argument("--language", default="ko")
    parser.add_argument("--context", default=None)
    args = parser.parse_args(argv)
    if args.serve and args.image is not None:
        parser.error("--serve and --image are mutually exclusive")
    if not args.serve and args.image is None:
        parser.error("one-shot mode requires --image; otherwise pass --serve")
    if args.min_pixels < 32 * 32:
        parser.error("--min-pixels must be at least 1024")
    if args.max_pixels < args.min_pixels:
        parser.error("--max-pixels must be greater than or equal to --min-pixels")
    if args.max_pixels > 4096 * 4096:
        parser.error("--max-pixels must not exceed 16777216")
    if args.max_new_tokens < 1 or args.max_new_tokens > 32_768:
        parser.error("--max-new-tokens must be between 1 and 32768")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        runtime = load_runtime(args)
    except Exception as exc:
        _log(f"Unable to start OvisOCR2 worker: {exc}")
        traceback.print_exc(file=sys.stderr)
        _emit(
            {
                "protocol": PROTOCOL,
                "type": "ready",
                "ok": False,
                "worker": "ovisocr2",
                "backend": BACKEND,
                "error": {
                    "type": type(exc).__name__,
                    "code": "startup_failed",
                    "message": str(exc),
                    "retryable": False,
                },
            }
        )
        return 1

    if args.serve:
        return serve(runtime)

    _emit(_ready_message(runtime))
    request = {
        "id": args.id,
        "image": str(args.image),
        "task": args.task,
        "language": args.language,
        "context": args.context,
    }
    _emit(_result_message(runtime, request))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
