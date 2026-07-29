"""Native Transformers worker for the pinned MonkeyOCRv2-B parser.

Protocol stdout is reserved for lines beginning with ``@@PPTX_WIKI@@``.
Everything else, including third-party model output, is redirected to stderr.

The OTSL conversion and repeat-output strategy are adapted from the official
MonkeyOCRv2 repository (Apache-2.0):
https://github.com/Yuliang-Liu/MonkeyOCRv2
"""

from __future__ import annotations

import argparse
import ast
import contextlib
import html
import json
import math
import os
import re
import sys
import tempfile
import traceback
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, TextIO

from download import verify_manifest
from model_spec import MODEL_REPO_ID, MODEL_REVISION

PROTOCOL_PREFIX = "@@PPTX_WIKI@@"
PROTOCOL = "pptx-wiki-ocr-worker/1"
BACKEND = "hf_monkeyocr_v2_b"
SUPPORTED_TASKS = {"document", "text", "table", "chart", "formula"}

TEXT_PROMPT = "Please output the text content from the image."
TABLE_PROMPT = "Please extract the table from the image and represent it in OTSL format."
FORMULA_PROMPT = (
    "Please write out the expression of the formula in the image using LaTeX format."
)
END2END_PROMPT = (
    "List the document elements in reading order, including their categories, "
    "coordinates, and the content of each element."
)

PROMPTS = {
    "document": END2END_PROMPT,
    "chart": END2END_PROMPT,
    "text": TEXT_PROMPT,
    "table": TABLE_PROMPT,
    "formula": FORMULA_PROMPT,
}


class WorkerError(RuntimeError):
    """A request or model-runtime error safe to return to the parent process."""


def _device_argument(value: str) -> str:
    normalized = value.strip().lower()
    if not re.fullmatch(r"(?:auto|cpu|cuda|gpu)(?::\d+)?", normalized):
        raise argparse.ArgumentTypeError(
            "device must be auto, cpu, cuda, cuda:N, gpu, or gpu:N"
        )
    if normalized.startswith("cpu:") or normalized.startswith("auto:"):
        raise argparse.ArgumentTypeError(
            "device must be auto, cpu, cuda, cuda:N, gpu, or gpu:N"
        )
    return normalized


def emit_protocol(payload: dict[str, Any], stream: TextIO = sys.stdout) -> None:
    framed = dict(payload)
    framed["protocol"] = PROTOCOL
    line = json.dumps(
        framed,
        ensure_ascii=False,
        allow_nan=False,
        separators=(",", ":"),
    )
    print(f"{PROTOCOL_PREFIX}{line}", file=stream, flush=True)


def log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary_name, path)
    except Exception:
        try:
            os.unlink(temporary_name)
        except OSError:
            pass
        raise


def _finite_bbox(value: Any) -> tuple[float, float, float, float] | None:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    try:
        numbers = tuple(float(item) for item in value)
    except (TypeError, ValueError):
        return None
    if not all(math.isfinite(item) for item in numbers):
        return None
    x1, y1, x2, y2 = numbers
    if x1 > x2:
        x1, x2 = x2, x1
    if y1 > y2:
        y1, y2 = y2, y1
    return x1, y1, x2, y2


def _normalized_bbox_to_pixels(
    value: Any, width: int, height: int
) -> list[float] | None:
    bbox = _finite_bbox(value)
    if bbox is None:
        return None
    x1, y1, x2, y2 = bbox
    x1 = max(0.0, min(float(width), x1 / 1000.0 * width))
    x2 = max(0.0, min(float(width), x2 / 1000.0 * width))
    y1 = max(0.0, min(float(height), y1 / 1000.0 * height))
    y2 = max(0.0, min(float(height), y2 / 1000.0 * height))
    if x2 < x1:
        x1, x2 = x2, x1
    if y2 < y1:
        y1, y2 = y2, y1
    return [x1, y1, x2, y2]


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = data.strip()
        if value:
            self.parts.append(value)


def html_to_text(value: str) -> str:
    parser = _TextExtractor()
    try:
        parser.feed(value)
    except Exception:
        return re.sub(r"<[^>]+>", " ", value).strip()
    return "\t".join(parser.parts)


class _TableSanitizer(HTMLParser):
    allowed_tags = {"table", "thead", "tbody", "tfoot", "tr", "td", "th", "caption"}
    allowed_attributes = {"rowspan", "colspan"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag not in self.allowed_tags:
            return
        safe_attributes = []
        for name, value in attrs:
            name = name.lower()
            if name not in self.allowed_attributes or value is None:
                continue
            try:
                number = max(1, min(100, int(value)))
            except ValueError:
                continue
            safe_attributes.append(f'{name}="{number}"')
        suffix = f" {' '.join(safe_attributes)}" if safe_attributes else ""
        self.parts.append(f"<{tag}{suffix}>")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in self.allowed_tags:
            self.parts.append(f"</{tag}>")

    def handle_data(self, data: str) -> None:
        self.parts.append(html.escape(data))


def sanitize_table_html(value: str) -> str:
    parser = _TableSanitizer()
    try:
        parser.feed(value)
        parser.close()
    except Exception:
        return "<table></table>"
    sanitized = "".join(parser.parts).strip()
    return sanitized if "<table" in sanitized.lower() else "<table></table>"


def otsl_to_html(otsl: str) -> str:
    """Convert MonkeyOCR's OTSL cell tokens to a rowspan/colspan HTML table."""

    if not otsl or not otsl.strip():
        return "<table></table>"
    row_strings = otsl.split("<nl>")
    if row_strings and not row_strings[-1].strip():
        row_strings.pop()
    grid: list[list[dict[str, Any] | None]] = []
    for row_index, row_string in enumerate(row_strings):
        while len(grid) <= row_index:
            grid.append([])
        parts = re.findall(r"<([a-z]+)>(.*?)(?=<[a-z]+>|$)", row_string, flags=re.DOTALL)
        column_index = 0
        for tag, content in parts:
            while True:
                while len(grid[row_index]) <= column_index:
                    grid[row_index].append(None)
                if grid[row_index][column_index] is None:
                    break
                column_index += 1
            if tag in {"fcel", "ecel"}:
                grid[row_index][column_index] = {
                    "text": content.strip() if tag == "fcel" else "",
                    "rowspan": 1,
                    "colspan": 1,
                    "valid": True,
                }
            elif tag == "lcel":
                source_column = column_index - 1
                source = None
                while source_column >= 0:
                    candidate = grid[row_index][source_column]
                    if candidate and candidate.get("valid"):
                        source = candidate
                        break
                    source_column -= 1
                if source is not None:
                    source["colspan"] += 1
                    grid[row_index][column_index] = {"valid": False, "type": "lcel"}
                else:
                    grid[row_index][column_index] = {
                        "text": "",
                        "rowspan": 1,
                        "colspan": 1,
                        "valid": True,
                    }
            elif tag == "ucel":
                source_row = row_index - 1
                source = None
                while source_row >= 0:
                    if len(grid[source_row]) > column_index:
                        candidate = grid[source_row][column_index]
                        if candidate and candidate.get("valid"):
                            source = candidate
                            break
                    source_row -= 1
                if source is not None:
                    source["rowspan"] += 1
                    grid[row_index][column_index] = {"valid": False, "type": "ucel"}
                else:
                    grid[row_index][column_index] = {
                        "text": "",
                        "rowspan": 1,
                        "colspan": 1,
                        "valid": True,
                    }
            elif tag == "xcel":
                grid[row_index][column_index] = {"valid": False, "type": "xcel"}
            column_index += 1

    output = ["<table>"]
    for row in grid:
        output.append("<tr>")
        for cell in row:
            if not cell or not cell.get("valid"):
                continue
            attributes = []
            if cell["rowspan"] > 1:
                attributes.append(f'rowspan="{int(cell["rowspan"])}"')
            if cell["colspan"] > 1:
                attributes.append(f'colspan="{int(cell["colspan"])}"')
            attribute_text = f" {' '.join(attributes)}" if attributes else ""
            output.append(f"<td{attribute_text}>{html.escape(str(cell['text']))}</td>")
        output.append("</tr>")
    output.append("</table>")
    return "".join(output)


def _strip_code_fence(value: str) -> str:
    stripped = value.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```[^\n]*\n?", "", stripped)
        stripped = re.sub(r"\n?```\s*$", "", stripped)
    return stripped.strip()


def _parse_document_items(raw: str) -> list[dict[str, Any]]:
    stripped = _strip_code_fence(raw)
    candidates = [stripped]
    first = stripped.find("[")
    last = stripped.rfind("]")
    if 0 <= first < last:
        candidates.append(stripped[first : last + 1])
    parsed: Any = None
    for candidate in candidates:
        try:
            parsed = ast.literal_eval(candidate)
        except (SyntaxError, ValueError):
            continue
        if isinstance(parsed, list):
            break
    if not isinstance(parsed, list):
        return []
    items = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        bbox = _finite_bbox(item.get("bbox"))
        label = item.get("label")
        if bbox is None or label is None:
            continue
        content = item.get("content", "")
        items.append(
            {
                "bbox": list(bbox),
                "label": str(label),
                "content": "" if content is None else str(content),
            }
        )
    return items


def detect_repeat_output(
    text: str,
    *,
    base_max_repeats: int = 4,
    window_size: int = 500,
    scaling_factor: float = 3.0,
) -> bool:
    if not text:
        return False
    for sequence_length in range(1, min(window_size // 2, len(text)) + 1):
        candidate = text[-sequence_length:]
        maximum = int(base_max_repeats * (1 + scaling_factor / sequence_length))
        repeats = 0
        position = len(text) - sequence_length
        while position >= 0 and text[position : position + sequence_length] == candidate:
            repeats += 1
            position -= sequence_length
        if repeats > maximum:
            return True
    return False


def _safe_context(value: Any, maximum: int = 2_000) -> str:
    if value is None:
        return ""
    text = str(value).replace("\x00", " ").strip()
    return text[:maximum]


def build_prompt(task: str, language: str, context: str) -> str:
    prompt = PROMPTS[task]
    hints = []
    if language:
        hints.append(f"Document language hint: {language}.")
    if context:
        hints.append(
            "Context hint for disambiguation only; never output text that is not visible in the image: "
            + context
        )
    return prompt if not hints else prompt + "\n" + "\n".join(hints)


def _formula_markdown(raw: str) -> str:
    content = raw.strip().strip("$").strip()
    return f"$$\n{content}\n$$" if content else ""


def _block_from_content(
    *,
    label: str,
    content: str,
    bbox: list[float],
    order: int,
    language: str,
) -> dict[str, Any]:
    normalized_label = label.strip() or "Text"
    lower = normalized_label.lower()
    block_html: str | None = None
    if lower == "table":
        block_html = (
            sanitize_table_html(content) if "<table" in content.lower() else otsl_to_html(content)
        )
        block_text = html_to_text(block_html)
        markdown = block_html
        kind = "table"
    elif lower in {"formula", "equation", "equation-block"}:
        markdown = _formula_markdown(content)
        block_text = content.strip()
        kind = "formula"
    elif lower == "title":
        block_text = content.strip()
        markdown = "# " + block_text.replace("\n", "\n# ") if block_text else ""
        kind = "title"
    elif lower in {"section-header", "section_header"}:
        block_text = content.strip()
        markdown = "## " + block_text.replace("\n", "\n## ") if block_text else ""
        kind = "section_header"
    else:
        block_text = content.strip()
        markdown = block_text
        kind = lower.replace("-", "_") or "text"
    return {
        "kind": kind,
        "text": block_text,
        "markdown": markdown,
        "html": block_html,
        "bbox": bbox,
        "confidence": None,
        "order": order,
        "metadata": {
            "raw_label": normalized_label,
            "bbox_scale": 1000,
            "language_hint": language,
            "model_revision": MODEL_REVISION,
        },
    }


def normalize_result(
    *,
    raw: str,
    task: str,
    language: str,
    width: int,
    height: int,
    warnings: list[str],
) -> dict[str, Any]:
    blocks: list[dict[str, Any]] = []
    if task in {"document", "chart"}:
        for order, item in enumerate(_parse_document_items(raw)):
            bbox = _normalized_bbox_to_pixels(item["bbox"], width, height)
            if bbox is None:
                continue
            blocks.append(
                _block_from_content(
                    label=item["label"],
                    content=item["content"],
                    bbox=bbox,
                    order=order,
                    language=language,
                )
            )
        if not blocks and raw.strip():
            warnings.append("END2END output was not a valid element list; preserved as plain text.")
            blocks.append(
                _block_from_content(
                    label="Chart" if task == "chart" else "Text",
                    content=raw,
                    bbox=[0.0, 0.0, float(width), float(height)],
                    order=0,
                    language=language,
                )
            )
    else:
        label = {"text": "Text", "table": "Table", "formula": "Formula"}[task]
        blocks.append(
            _block_from_content(
                label=label,
                content=raw,
                bbox=[0.0, 0.0, float(width), float(height)],
                order=0,
                language=language,
            )
        )

    text = "\n".join(block["text"] for block in blocks if block["text"]).strip()
    markdown = "\n\n".join(
        block["markdown"] for block in blocks if isinstance(block.get("markdown"), str) and block["markdown"]
    ).strip()
    table_html = [block["html"] for block in blocks if isinstance(block.get("html"), str)]
    return {
        "text": text,
        "markdown": markdown,
        "html": "\n".join(table_html) if table_html else None,
        "blocks": blocks,
        "confidence": None,
        "raw": raw,
        "warnings": warnings,
        "metadata": {
            "model": MODEL_REPO_ID,
            "revision": MODEL_REVISION,
            "task": task,
        },
    }


@dataclass(slots=True)
class Runtime:
    model_dir: Path
    device_name: str
    dtype_name: str
    max_new_tokens: int
    repeat_retries: int
    torch: Any
    processor: Any
    model: Any
    input_device: Any
    model_dtype: Any
    startup_warnings: list[str] = field(default_factory=list)

    @classmethod
    def load(
        cls,
        *,
        model_dir: Path,
        device_name: str,
        dtype_name: str,
        max_new_tokens: int,
        repeat_retries: int,
    ) -> "Runtime":
        errors = verify_manifest(model_dir)
        if errors:
            joined = "\n".join(f"  - {error}" for error in errors[:20])
            raise WorkerError(f"Model integrity verification failed:\n{joined}")

        # No Hub or telemetry access is permitted after the explicit download step.
        os.environ["HF_HUB_OFFLINE"] = "1"
        os.environ["TRANSFORMERS_OFFLINE"] = "1"
        os.environ["HF_HUB_DISABLE_TELEMETRY"] = "1"
        os.environ["TOKENIZERS_PARALLELISM"] = "false"

        with contextlib.redirect_stdout(sys.stderr):
            try:
                import torch
                from transformers import AutoConfig, AutoModelForCausalLM, AutoProcessor
            except ImportError as exc:
                raise WorkerError(
                    "Worker dependencies are missing; run setup-windows.ps1 first."
                ) from exc

            if device_name == "auto":
                device_name = "cuda:0" if torch.cuda.is_available() else "cpu"
            elif device_name == "gpu":
                device_name = "cuda:0"
            elif device_name.startswith("gpu:"):
                device_name = "cuda:" + device_name.split(":", 1)[1]
            elif device_name == "cuda":
                device_name = "cuda:0"

            torch_device = torch.device(device_name)
            if torch_device.type == "cuda":
                if not torch.cuda.is_available():
                    raise WorkerError(
                        f"--device {device_name} was requested, but PyTorch cannot see a CUDA GPU."
                    )
                device_index = torch_device.index if torch_device.index is not None else 0
                if device_index >= torch.cuda.device_count():
                    raise WorkerError(f"CUDA device index {device_index} is unavailable.")
                torch.cuda.set_device(device_index)

            startup_warnings: list[str] = []
            if dtype_name == "auto":
                if torch_device.type == "cuda":
                    dtype_name = "bfloat16" if torch.cuda.is_bf16_supported() else "float16"
                else:
                    dtype_name = "float32"
            dtype_map = {
                "bfloat16": torch.bfloat16,
                "float16": torch.float16,
                "float32": torch.float32,
            }
            model_dtype = dtype_map[dtype_name]
            if torch_device.type == "cpu" and model_dtype == torch.float16:
                raise WorkerError("float16 CPU inference is unsupported; use --dtype float32 or auto.")
            if (
                torch_device.type == "cuda"
                and model_dtype == torch.bfloat16
                and not torch.cuda.is_bf16_supported()
            ):
                raise WorkerError("This GPU does not support bfloat16; use --dtype float16 or auto.")

            local = str(model_dir)
            config = AutoConfig.from_pretrained(
                local,
                trust_remote_code=True,
                local_files_only=True,
            )
            config.use_cache = True
            config._attn_implementation = "sdpa"
            if hasattr(config, "vision_config"):
                # The pinned custom SDPA vision implementation explicitly
                # selects EFFICIENT_ATTENTION, which has no CPU kernel in
                # torch 2.6.  Its eager implementation is the supported CPU
                # fallback; CUDA keeps the faster upstream SDPA path.
                config.vision_config.attn_implementation = (
                    "eager" if torch_device.type == "cpu" else "sdpa"
                )

            processor = AutoProcessor.from_pretrained(
                local,
                trust_remote_code=True,
                local_files_only=True,
            )
            model = AutoModelForCausalLM.from_pretrained(
                local,
                config=config,
                trust_remote_code=True,
                local_files_only=True,
                dtype=model_dtype,
                device_map={"": device_name},
                low_cpu_mem_usage=True,
            ).eval()

            # The upstream vision forward forces BF16 inputs.  Keep that exact path on
            # supported GPUs; for CPU/Turing compatibility, use the loaded weight dtype.
            if model_dtype != torch.bfloat16:
                original_vision_forward = model.vision_tower.forward
                vision_tower = model.vision_tower

                def compatible_vision_forward(hidden_states, grid_thw, bf16=True):
                    del bf16
                    return original_vision_forward(
                        hidden_states.to(dtype=vision_tower.dtype), grid_thw, bf16=False
                    )

                model.vision_tower.forward = compatible_vision_forward
                startup_warnings.append(
                    "Non-BF16 compatibility mode is active; this differs slightly from the official BF16 path."
                )
            if torch_device.type == "cpu":
                startup_warnings.append(
                    "CPU inference uses eager vision attention and can be extremely slow "
                    "for long OCR generations."
                )

            model.generation_config.do_sample = False
            input_device = next(model.parameters()).device

        return cls(
            model_dir=model_dir,
            device_name=device_name,
            dtype_name=dtype_name,
            max_new_tokens=max_new_tokens,
            repeat_retries=repeat_retries,
            torch=torch,
            processor=processor,
            model=model,
            input_device=input_device,
            model_dtype=model_dtype,
            startup_warnings=startup_warnings,
        )

    def ready_payload(self) -> dict[str, Any]:
        return {
            "type": "ready",
            "protocol": PROTOCOL,
            "ok": True,
            "backend": BACKEND,
            "model": MODEL_REPO_ID,
            "revision": MODEL_REVISION,
            "device": self.device_name,
            "dtype": self.dtype_name,
            "pid": os.getpid(),
            "warnings": self.startup_warnings,
        }

    def _generate_once(
        self,
        *,
        image: Any,
        prompt: str,
        do_sample: bool,
        temperature: float | None = None,
    ) -> str:
        messages = [
            {"role": "system", "content": "You are a helpful assistant."},
            {
                "role": "user",
                "content": [
                    {"type": "image"},
                    {"type": "text", "text": prompt},
                ],
            },
        ]
        rendered = self.processor.apply_chat_template(
            messages,
            tokenize=False,
            add_generation_prompt=True,
        )
        inputs = self.processor(
            text=[rendered],
            images=[image],
            padding=True,
            return_tensors="pt",
        )
        for key, value in list(inputs.items()):
            if self.torch.is_tensor(value):
                inputs[key] = value.to(self.input_device)
        generation_args: dict[str, Any] = {
            "max_new_tokens": self.max_new_tokens,
            "do_sample": do_sample,
            "use_cache": True,
        }
        if do_sample:
            generation_args.update(temperature=temperature or 0.2, top_p=0.95)
        with self.torch.inference_mode(), contextlib.redirect_stdout(sys.stderr):
            generated = self.model.generate(**inputs, **generation_args)
        prompt_length = inputs["input_ids"].shape[1]
        generated = generated[:, prompt_length:]
        return self.processor.batch_decode(
            generated,
            skip_special_tokens=True,
            clean_up_tokenization_spaces=False,
        )[0].strip()

    def recognize(self, request: dict[str, Any]) -> dict[str, Any]:
        request_id = request.get("id")
        if (
            not isinstance(request_id, (str, int))
            or isinstance(request_id, bool)
            or (isinstance(request_id, str) and not request_id.strip())
        ):
            raise WorkerError("request id must be a non-empty string or integer")
        supplied_protocol = request.get("protocol")
        if supplied_protocol is not None and supplied_protocol != PROTOCOL:
            raise WorkerError(f"unsupported protocol: {supplied_protocol!r}")
        task = str(request.get("task", "document")).lower()
        if task not in SUPPORTED_TASKS:
            raise WorkerError(f"unsupported task: {task!r}")
        image_value = request.get("image")
        image_path_value = request.get("image_path")
        if (
            image_value is not None
            and image_path_value is not None
            and image_value != image_path_value
        ):
            raise WorkerError("image and image_path must match when both are provided")
        if image_value is None:
            image_value = image_path_value
        if not isinstance(image_value, str) or not image_value.strip():
            raise WorkerError("image or image_path must be a non-empty local file path")
        if re.match(r"^[a-z][a-z0-9+.-]*://", image_value, flags=re.IGNORECASE):
            raise WorkerError("image URLs are not accepted; provide a downloaded local file")
        image_path = Path(image_value).expanduser().resolve()
        if not image_path.is_file():
            raise WorkerError(f"image does not exist: {image_path}")
        language = _safe_context(request.get("language", "ko"), 64)
        context = _safe_context(request.get("context"), 2_000)
        prompt = build_prompt(task, language, context)
        warnings = list(self.startup_warnings)

        try:
            from PIL import Image, ImageOps
        except ImportError as exc:  # pragma: no cover - setup installs Pillow
            raise WorkerError("Pillow is not installed") from exc
        try:
            with Image.open(image_path) as opened:
                opened.load()
                image = ImageOps.exif_transpose(opened).convert("RGB")
        except Exception as exc:
            raise WorkerError(f"unable to open image: {exc}") from exc

        width, height = image.size
        if width < 1 or height < 1:
            raise WorkerError("image has invalid dimensions")
        if width * height > 100_000_000:
            raise WorkerError("image exceeds the 100-megapixel safety limit")

        try:
            raw = self._generate_once(image=image, prompt=prompt, do_sample=False)
            retries = 0
            while detect_repeat_output(raw) and retries < self.repeat_retries:
                retries += 1
                raw = self._generate_once(
                    image=image,
                    prompt=prompt,
                    do_sample=True,
                    temperature=min(0.2 * retries, 0.8),
                )
        finally:
            image.close()
        if retries:
            warnings.append(f"Repeated decoder output was regenerated {retries} time(s).")
        if detect_repeat_output(raw):
            warnings.append("Decoder output still appears repetitive after retries.")
        result = normalize_result(
            raw=raw,
            task=task,
            language=language,
            width=width,
            height=height,
            warnings=warnings,
        )
        result["metadata"]["request_id"] = str(request_id)
        result["metadata"]["context_supplied"] = bool(context)
        return result


def _error_payload(request_id: Any, exc: Exception) -> dict[str, Any]:
    return {
        "type": "result",
        "protocol": PROTOCOL,
        "id": request_id,
        "ok": False,
        "error": {"type": type(exc).__name__, "message": str(exc)},
    }


def serve(runtime: Runtime) -> int:
    emit_protocol(runtime.ready_payload())
    for line_number, line in enumerate(sys.stdin, start=1):
        if len(line) > 1_000_000:
            emit_protocol(_error_payload(None, WorkerError("request line exceeds 1 MB")))
            continue
        if not line.strip():
            continue
        request_id = None
        try:
            request = json.loads(line)
            if not isinstance(request, dict):
                raise WorkerError("each JSONL request must be an object")
            request_id = request.get("id")
            supplied_protocol = request.get("protocol")
            if supplied_protocol is not None and supplied_protocol != PROTOCOL:
                raise WorkerError(f"unsupported protocol: {supplied_protocol!r}")
            message_type = request.get("type")
            operation = {
                "request": request.get("op", "recognize"),
                "shutdown": "shutdown",
                "ping": "ping",
            }.get(message_type, request.get("op", "recognize"))
            if operation == "shutdown":
                emit_protocol(
                    {
                        "type": "result",
                        "protocol": PROTOCOL,
                        "ok": True,
                        "id": request.get("id"),
                        "shutdown": True,
                    }
                )
                return 0
            if operation == "ping":
                emit_protocol(
                    {
                        "type": "result",
                        "protocol": PROTOCOL,
                        "ok": True,
                        "id": request.get("id"),
                        "pong": True,
                    }
                )
                continue
            if operation != "recognize":
                raise WorkerError(f"unsupported operation: {operation!r}")
            result = runtime.recognize(request)
            emit_protocol(
                {
                    "type": "result",
                    "protocol": PROTOCOL,
                    "id": request_id,
                    "ok": True,
                    "result": result,
                }
            )
        except Exception as exc:
            log(f"Request failed on JSONL line {line_number}: {exc}")
            emit_protocol(_error_payload(request_id, exc))
    return 0


def parse_args() -> argparse.Namespace:
    worker_root = Path(__file__).resolve().parent
    repo_root = worker_root.parent.parent
    parser = argparse.ArgumentParser(description="Run the local MonkeyOCRv2-B worker.")
    parser.add_argument("--serve", action="store_true", help="Read persistent requests from stdin as JSONL.")
    parser.add_argument(
        "--model-dir",
        type=Path,
        default=repo_root / "models" / "monkeyocr_v2_b",
    )
    parser.add_argument(
        "--device",
        type=_device_argument,
        default="auto",
        help="auto, cpu, cuda, cuda:N, gpu, or gpu:N",
    )
    parser.add_argument(
        "--dtype",
        choices=["auto", "bfloat16", "float16", "float32"],
        default="auto",
    )
    parser.add_argument("--max-new-tokens", type=int, default=16_384)
    parser.add_argument("--repeat-retries", type=int, default=3)

    # One-shot request fields. Persistent mode receives the same fields via JSONL.
    parser.add_argument("--request-id", default="one-shot")
    parser.add_argument("--image", type=str)
    parser.add_argument("--task", choices=sorted(SUPPORTED_TASKS), default="document")
    parser.add_argument("--language", default="ko")
    parser.add_argument("--context", default="")
    parser.add_argument("--output", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.serve and not args.image:
        print("--image is required unless --serve is used", file=sys.stderr)
        return 2

    model_dir = args.model_dir.expanduser().resolve()
    try:
        if not 1 <= args.max_new_tokens <= 32_768:
            raise WorkerError("--max-new-tokens must be between 1 and 32768")
        if not 0 <= args.repeat_retries <= 10:
            raise WorkerError("--repeat-retries must be between 0 and 10")
        runtime = Runtime.load(
            model_dir=model_dir,
            device_name=args.device,
            dtype_name=args.dtype,
            max_new_tokens=args.max_new_tokens,
            repeat_retries=args.repeat_retries,
        )
    except Exception as exc:
        log(f"Unable to initialize MonkeyOCRv2: {exc}")
        traceback.print_exc(file=sys.stderr)
        emit_protocol(
            {
                "type": "ready",
                "protocol": PROTOCOL,
                "ok": False,
                "backend": BACKEND,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            }
        )
        return 2

    if args.serve:
        return serve(runtime)

    emit_protocol(runtime.ready_payload())
    request = {
        "id": args.request_id,
        "image": args.image,
        "task": args.task,
        "language": args.language,
        "context": args.context,
    }
    try:
        result = runtime.recognize(request)
    except Exception as exc:
        log(f"One-shot request failed: {exc}")
        traceback.print_exc(file=sys.stderr)
        emit_protocol(_error_payload(args.request_id, exc))
        return 1
    if args.output:
        _atomic_write_json(args.output.expanduser().resolve(), result)
    emit_protocol(
        {
            "type": "result",
            "protocol": PROTOCOL,
            "id": args.request_id,
            "ok": True,
            "result": result,
        }
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
