"""OCR adapters with a small, dependency-light interchange format.

The rest of the project deliberately does not import vendor-specific OCR types.
An :class:`OCRRequest` contains either an already isolated image ROI or a source
image plus a pixel crop.  Every adapter returns the same :class:`OCRResult`.

The Paddle adapter invokes the *full* ``paddleocr doc_parser`` pipeline.  This is
intentional: directly serving the PaddleOCR-VL checkpoint omits the layout stage
and is not equivalent to the document parser.  For an object-level ROI, layout
detection is disabled and the element prompt is selected explicitly, preventing
two neighbouring PowerPoint tables from being merged a second time by OCR.
"""

from __future__ import annotations

import base64
from collections import deque
import json
import math
import os
import queue
import re
import subprocess
import tempfile
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Callable, Literal, Mapping, Protocol, Sequence
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

OCRTask = Literal["document", "text", "table", "chart", "formula"]
PixelBBox = tuple[float, float, float, float]
ImageSource = str | Path | bytes


class OCRError(RuntimeError):
    """Base class for adapter errors."""


class OCRConfigurationError(OCRError):
    """An adapter is not usable with the supplied configuration."""


class OCRExecutionError(OCRError):
    """An OCR process or endpoint failed."""


class OCRTransientError(OCRExecutionError):
    """An execution error which is reasonable to retry."""


class OCRParseError(OCRError):
    """A backend response could not be converted to the common schema."""


@dataclass(frozen=True, slots=True)
class RetryPolicy:
    attempts: int = 3
    initial_delay: float = 0.5
    multiplier: float = 2.0
    max_delay: float = 4.0

    def __post_init__(self) -> None:
        if self.attempts < 1:
            raise ValueError("attempts must be at least 1")
        if min(self.initial_delay, self.multiplier, self.max_delay) < 0:
            raise ValueError("retry delays and multiplier must be non-negative")


@dataclass(frozen=True, slots=True)
class OCRRequest:
    """One OCR unit.

    ``crop`` uses half-open pixel coordinates ``(left, top, right, bottom)`` and
    is applied before a backend sees the image.  Backend bounding boxes are thus
    always relative to the cropped ROI.  ``source_bbox`` can retain the ROI's
    coordinate in the rendered slide for the caller; adapters do not interpret
    it.
    """

    image: ImageSource
    task: OCRTask = "document"
    crop: tuple[int, int, int, int] | None = None
    source_bbox: PixelBBox | None = None
    language_hint: str = "ko"
    context: str | None = None
    request_id: str | None = None
    media_type: str | None = None

    def __post_init__(self) -> None:
        if self.task not in {"document", "text", "table", "chart", "formula"}:
            raise ValueError(f"unsupported OCR task: {self.task!r}")
        if self.crop is not None:
            left, top, right, bottom = self.crop
            if min(left, top) < 0 or right <= left or bottom <= top:
                raise ValueError("crop must be a non-empty, non-negative pixel box")


@dataclass(slots=True)
class OCRBlock:
    kind: str
    text: str = ""
    markdown: str | None = None
    html: str | None = None
    bbox: PixelBBox | None = None
    confidence: float | None = None
    confidence_source: str | None = None
    order: int | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class OCRResult:
    backend: str
    text: str = ""
    markdown: str = ""
    html: str | None = None
    blocks: list[OCRBlock] = field(default_factory=list)
    confidence: float | None = None
    raw: Any = None
    warnings: list[str] = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return not (self.text.strip() or self.markdown.strip() or (self.html or "").strip())

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class OCRAdapter(Protocol):
    name: str

    def recognize(self, request: OCRRequest) -> OCRResult: ...


@dataclass(slots=True)
class _MaterializedImage:
    path: Path
    media_type: str
    temporary_dir: tempfile.TemporaryDirectory[str] | None = None

    def close(self) -> None:
        if self.temporary_dir is not None:
            self.temporary_dir.cleanup()

    def __enter__(self) -> "_MaterializedImage":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()


def _guess_media_type(data: bytes, path: Path | None = None) -> str:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return "image/webp"
    if path is not None:
        suffix = path.suffix.lower()
        return {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
            ".gif": "image/gif",
            ".bmp": "image/bmp",
            ".tif": "image/tiff",
            ".tiff": "image/tiff",
        }.get(suffix, "application/octet-stream")
    return "application/octet-stream"


def _materialize_image(request: OCRRequest) -> _MaterializedImage:
    source_path: Path | None = None
    data: bytes | None = None
    if isinstance(request.image, bytes):
        if not request.image:
            raise OCRConfigurationError("image bytes are empty")
        data = request.image
    else:
        source_path = Path(request.image).expanduser()
        if not source_path.is_file():
            raise OCRConfigurationError(f"image does not exist: {source_path}")

    if request.crop is None and source_path is not None:
        media_type = request.media_type or _guess_media_type(source_path.read_bytes()[:16], source_path)
        return _MaterializedImage(source_path.resolve(), media_type)

    if request.crop is None and data is not None:
        media_type = request.media_type or _guess_media_type(data[:16])
        suffix = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/webp": ".webp",
            "image/gif": ".gif",
        }.get(media_type, ".img")
        temp_dir = tempfile.TemporaryDirectory(prefix="pptx-wiki-ocr-")
        output_path = Path(temp_dir.name) / f"roi{suffix}"
        output_path.write_bytes(data)
        return _MaterializedImage(output_path, media_type, temp_dir)

    temp_dir = tempfile.TemporaryDirectory(prefix="pptx-wiki-ocr-")
    output_path = Path(temp_dir.name) / "roi.png"
    try:
        try:
            from PIL import Image
        except ImportError as exc:  # pragma: no cover - declared as a project dependency
            raise OCRConfigurationError("Pillow is required to crop OCR image inputs") from exc
        if source_path is not None:
            image = Image.open(source_path)
        else:
            from io import BytesIO

            image = Image.open(BytesIO(data or b""))
        with image:
            image.load()
            if request.crop is not None:
                left, top, right, bottom = request.crop
                if right > image.width or bottom > image.height:
                    raise OCRConfigurationError(
                        f"crop {request.crop!r} exceeds image dimensions {(image.width, image.height)!r}"
                    )
                image = image.crop(request.crop)
            if image.mode not in {"RGB", "RGBA", "L"}:
                image = image.convert("RGB")
            image.save(output_path, format="PNG")
    except OCRError:
        temp_dir.cleanup()
        raise
    except Exception as exc:
        temp_dir.cleanup()
        raise OCRConfigurationError(f"unable to read OCR image: {exc}") from exc
    return _MaterializedImage(output_path, "image/png", temp_dir)


def _retry(operation: Callable[[], OCRResult], policy: RetryPolicy) -> OCRResult:
    delay = policy.initial_delay
    for attempt in range(1, policy.attempts + 1):
        try:
            return operation()
        except OCRTransientError:
            if attempt >= policy.attempts:
                raise
            if delay:
                time.sleep(min(delay, policy.max_delay))
            delay = min(policy.max_delay, delay * policy.multiplier)
    raise AssertionError("unreachable")


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = data.strip()
        if value:
            self.parts.append(value)


def _html_to_text(value: str) -> str:
    parser = _TextExtractor()
    try:
        parser.feed(value)
    except Exception:
        return re.sub(r"<[^>]+>", " ", value).strip()
    return "\t".join(parser.parts)


_HTML_TABLE_PATTERN = re.compile(r"<table\b[^>]*>.*?</table\s*>", re.IGNORECASE | re.DOTALL)


def _markdown_blocks(value: str) -> list[OCRBlock]:
    """Preserve multiple HTML tables as distinct blocks in Markdown output."""

    blocks: list[OCRBlock] = []
    cursor = 0
    for match in _HTML_TABLE_PATTERN.finditer(value):
        prefix = value[cursor : match.start()].strip()
        if prefix:
            blocks.append(
                OCRBlock(kind="text", text=_html_to_text(prefix), markdown=prefix, order=len(blocks))
            )
        table = match.group(0).strip()
        blocks.append(
            OCRBlock(
                kind="table",
                text=_html_to_text(table),
                markdown=table,
                html=table,
                order=len(blocks),
            )
        )
        cursor = match.end()
    suffix = value[cursor:].strip()
    if suffix:
        blocks.append(
            OCRBlock(kind="text", text=_html_to_text(suffix), markdown=suffix, order=len(blocks))
        )
    return blocks


def _clean_confidence(value: Any) -> float | None:
    try:
        confidence = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(confidence) or confidence < 0 or confidence > 1:
        return None
    return confidence


def _clean_bbox(value: Any) -> PixelBBox | None:
    if isinstance(value, Mapping):
        if all(key in value for key in ("x1", "y1", "x2", "y2")):
            value = [value["x1"], value["y1"], value["x2"], value["y2"]]
        elif all(key in value for key in ("left", "top", "right", "bottom")):
            value = [value["left"], value["top"], value["right"], value["bottom"]]
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)) or len(value) != 4:
        return None
    try:
        bbox = tuple(float(item) for item in value)
    except (TypeError, ValueError):
        return None
    if not all(math.isfinite(item) for item in bbox):
        return None
    x1, y1, x2, y2 = bbox
    if x2 < x1 or y2 < y1:
        return None
    return x1, y1, x2, y2


def _iou(first: PixelBBox | None, second: PixelBBox | None) -> float:
    if first is None or second is None:
        return 0.0
    ax1, ay1, ax2, ay2 = first
    bx1, by1, bx2, by2 = second
    intersection = max(0.0, min(ax2, bx2) - max(ax1, bx1)) * max(0.0, min(ay2, by2) - max(ay1, by1))
    union = max(0.0, ax2 - ax1) * max(0.0, ay2 - ay1) + max(0.0, bx2 - bx1) * max(
        0.0, by2 - by1
    ) - intersection
    return intersection / union if union else 0.0


def _extract_json(text: str) -> Any:
    """Parse JSON even when a chat model wraps it in prose or a code fence."""

    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped, flags=re.IGNORECASE)
        stripped = re.sub(r"\s*```$", "", stripped)
    try:
        return json.loads(stripped)
    except json.JSONDecodeError as first_error:
        decoder = json.JSONDecoder()
        for index, character in enumerate(stripped):
            if character not in "[{":
                continue
            try:
                value, _ = decoder.raw_decode(stripped[index:])
                return value
            except json.JSONDecodeError:
                continue
        raise OCRParseError(f"response is not valid JSON: {first_error}") from first_error


def _unwrap_payload(payload: Any) -> Any:
    if isinstance(payload, Mapping) and isinstance(payload.get("res"), Mapping):
        return payload["res"]
    return payload


def _paddle_layout_candidates(payload: Mapping[str, Any]) -> list[dict[str, Any]]:
    layout = payload.get("layout_det_res")
    if not isinstance(layout, Mapping):
        return []
    boxes = layout.get("boxes")
    return [dict(item) for item in boxes if isinstance(item, Mapping)] if isinstance(boxes, list) else []


def _confidence_for_paddle_block(
    kind: str, bbox: PixelBBox | None, candidates: Sequence[Mapping[str, Any]]
) -> float | None:
    scored: list[tuple[float, float]] = []
    for candidate in candidates:
        candidate_bbox = _clean_bbox(candidate.get("coordinate") or candidate.get("bbox"))
        overlap = _iou(bbox, candidate_bbox)
        label = str(candidate.get("label", "")).lower()
        label_bonus = 0.05 if label == kind.lower() else 0.0
        score = _clean_confidence(candidate.get("score"))
        if score is not None:
            scored.append((overlap + label_bonus, score))
    if not scored:
        return None
    match_quality, confidence = max(scored, key=lambda item: item[0])
    return confidence if match_quality >= 0.5 else None


def normalize_ocr_payload(
    payload: Any,
    *,
    backend: str,
    markdown_override: str | None = None,
    warnings: Sequence[str] = (),
) -> OCRResult:
    """Convert Paddle, command, or prompted-VLM JSON to :class:`OCRResult`."""

    raw = payload
    payload = _unwrap_payload(payload)
    if not isinstance(payload, Mapping):
        raise OCRParseError(f"{backend} response must be a JSON object")

    blocks: list[OCRBlock] = []
    paddle_blocks = payload.get("parsing_res_list")
    is_paddle = isinstance(paddle_blocks, list)
    source_blocks = paddle_blocks if is_paddle else payload.get("blocks", [])
    if source_blocks is None:
        source_blocks = []
    if not isinstance(source_blocks, list):
        raise OCRParseError("blocks must be a JSON array")

    layout_candidates = _paddle_layout_candidates(payload) if is_paddle else []
    for index, item in enumerate(source_blocks):
        if not isinstance(item, Mapping):
            continue
        kind = str(item.get("block_label") or item.get("kind") or item.get("type") or "text")
        content = str(item.get("block_content") or item.get("content") or item.get("text") or "").strip()
        block_html = item.get("html") if isinstance(item.get("html"), str) else None
        if block_html is None and kind.lower() == "table" and "<table" in content.lower():
            block_html = content
        block_markdown = item.get("markdown") if isinstance(item.get("markdown"), str) else None
        if block_markdown is None and block_html is None:
            block_markdown = content
        block_text = str(item.get("text") or "").strip()
        if not block_text:
            block_text = _html_to_text(block_html) if block_html else content
        bbox = _clean_bbox(item.get("block_bbox") or item.get("bbox"))
        confidence_value = item.get("confidence") if "confidence" in item else item.get("score")
        confidence = _clean_confidence(confidence_value)
        confidence_source = "model_reported" if confidence is not None else None
        if is_paddle and confidence is None:
            confidence = _confidence_for_paddle_block(kind, bbox, layout_candidates)
            confidence_source = "layout_detector" if confidence is not None else None
        order_value = item.get("block_order", item.get("order", index))
        try:
            order = int(order_value) if order_value is not None else None
        except (TypeError, ValueError):
            order = index
        known = {
            "block_label",
            "kind",
            "type",
            "block_content",
            "content",
            "text",
            "html",
            "markdown",
            "block_bbox",
            "bbox",
            "confidence",
            "score",
            "block_order",
            "order",
        }
        blocks.append(
            OCRBlock(
                kind=kind,
                text=block_text,
                markdown=block_markdown,
                html=block_html,
                bbox=bbox,
                confidence=confidence,
                confidence_source=confidence_source,
                order=order,
                metadata={str(key): value for key, value in item.items() if key not in known},
            )
        )

    if not blocks:
        raw_markdown = payload.get("markdown")
        if isinstance(raw_markdown, str) and raw_markdown.strip():
            blocks = _markdown_blocks(raw_markdown.strip())

    top_text = payload.get("text")
    text_value = str(top_text).strip() if isinstance(top_text, str) else "\n".join(
        block.text for block in blocks if block.text
    )
    top_markdown = payload.get("markdown")
    if isinstance(top_markdown, Mapping):
        top_markdown = top_markdown.get("text") or top_markdown.get("markdown_texts")
    markdown_value = markdown_override
    if markdown_value is None:
        markdown_value = str(top_markdown).strip() if isinstance(top_markdown, str) else "\n\n".join(
            (block.markdown or block.html or block.text)
            for block in blocks
            if block.markdown or block.html or block.text
        )
    top_html = payload.get("html")
    html_value = str(top_html).strip() if isinstance(top_html, str) and top_html.strip() else None
    if html_value is None:
        table_html = [block.html for block in blocks if block.html]
        html_value = "\n".join(table_html) if table_html else None

    result_confidence = _clean_confidence(payload.get("confidence"))
    if result_confidence is None:
        available = [block.confidence for block in blocks if block.confidence is not None]
        result_confidence = sum(available) / len(available) if available else None

    result_warnings = list(warnings)
    payload_warnings = payload.get("warnings")
    if isinstance(payload_warnings, list):
        result_warnings.extend(str(item) for item in payload_warnings if isinstance(item, str))
    if is_paddle and any(block.confidence is None for block in blocks):
        result_warnings.append(
            "PaddleOCR-VL does not expose recognition confidence for every block; missing values remain null."
        )
    return OCRResult(
        backend=backend,
        text=text_value,
        markdown=markdown_value or "",
        html=html_value,
        blocks=blocks,
        confidence=result_confidence,
        raw=raw,
        warnings=result_warnings,
    )


Runner = Callable[..., subprocess.CompletedProcess[str]]


class PaddleOCRCLIAdapter:
    """Run the local PaddleOCR-VL v1.6 full document-parsing pipeline."""

    name = "paddle_cli"

    def __init__(
        self,
        *,
        executable: str = "paddleocr",
        pipeline_version: str = "v1.6",
        device: str | None = None,
        engine: str | None = None,
        timeout: float = 300.0,
        retry: RetryPolicy | None = None,
        extra_args: Sequence[str] = (),
        scrub_env_vars: Sequence[str] = (),
        runner: Runner = subprocess.run,
    ) -> None:
        self.executable = executable
        self.pipeline_version = pipeline_version
        self.device = device
        self.engine = engine
        self.timeout = timeout
        self.retry = retry or RetryPolicy()
        self.extra_args = tuple(extra_args)
        self.scrub_env_vars = tuple(dict.fromkeys(name for name in scrub_env_vars if name))
        self._runner = runner

    def _command(self, image: Path, output_dir: Path, request: OCRRequest) -> list[str]:
        command = [
            self.executable,
            "doc_parser",
            "-i",
            str(image),
            "--pipeline_version",
            self.pipeline_version,
            "--save_path",
            str(output_dir),
            "--format_block_content",
            "True",
        ]
        if request.task == "document":
            # Retain close but distinct layout boxes.  Object-level ROI cropping is
            # still the primary guard against neighbouring-table fusion.
            command.extend(["--merge_layout_blocks", "False", "--layout_merge_bboxes_mode", "union"])
        else:
            prompt_label = "ocr" if request.task == "text" else request.task
            command.extend(["--use_layout_detection", "False", "--prompt_label", prompt_label])
        if self.device:
            command.extend(["--device", self.device])
        if self.engine:
            command.extend(["--engine", self.engine])
        command.extend(self.extra_args)
        return command

    def recognize(self, request: OCRRequest) -> OCRResult:
        with _materialize_image(request) as materialized:

            def attempt() -> OCRResult:
                with tempfile.TemporaryDirectory(prefix="pptx-wiki-paddle-") as output_name:
                    output_dir = Path(output_name)
                    command = self._command(materialized.path, output_dir, request)
                    child_env = os.environ.copy()
                    for name in self.scrub_env_vars:
                        child_env.pop(name, None)
                    try:
                        completed = self._runner(
                            command,
                            env=child_env,
                            capture_output=True,
                            text=True,
                            timeout=self.timeout,
                            check=False,
                        )
                    except FileNotFoundError as exc:
                        raise OCRConfigurationError(
                            f"PaddleOCR executable was not found: {self.executable!r}"
                        ) from exc
                    except subprocess.TimeoutExpired as exc:
                        raise OCRTransientError(f"PaddleOCR timed out after {self.timeout:g}s") from exc
                    except OSError as exc:
                        raise OCRTransientError(f"unable to launch PaddleOCR: {exc}") from exc
                    if completed.returncode != 0:
                        detail = (completed.stderr or completed.stdout or "").strip()[-2000:]
                        error_type = OCRExecutionError if completed.returncode == 2 else OCRTransientError
                        raise error_type(f"PaddleOCR exited with {completed.returncode}: {detail}")

                    json_files = sorted(output_dir.rglob("*_res.json")) or sorted(output_dir.rglob("*.json"))
                    if not json_files:
                        raise OCRTransientError("PaddleOCR completed without a JSON result")
                    try:
                        payload = json.loads(json_files[0].read_text(encoding="utf-8"))
                    except (OSError, json.JSONDecodeError) as exc:
                        raise OCRTransientError(f"invalid PaddleOCR JSON output: {exc}") from exc
                    markdown_files = sorted(output_dir.rglob("*.md"))
                    markdown = markdown_files[0].read_text(encoding="utf-8") if markdown_files else None
                    return normalize_ocr_payload(payload, backend=self.name, markdown_override=markdown)

            return _retry(attempt, self.retry)


class CommandOCRAdapter:
    """Invoke a Hugging Face/custom inference script without importing it.

    Command items support ``{image}``, ``{output}``, ``{task}``, ``{language}``,
    and ``{request_id}`` placeholders.  The child process must either write JSON
    to ``{output}`` or emit it on stdout.  Plain Markdown stdout is accepted as a
    compatibility fallback.  The command is always run with ``shell=False``.
    """

    name = "command"

    def __init__(
        self,
        command: Sequence[str],
        *,
        backend_name: str = "command",
        cwd: str | Path | None = None,
        env: Mapping[str, str] | None = None,
        scrub_env_vars: Sequence[str] = (),
        timeout: float = 300.0,
        retry: RetryPolicy | None = None,
        runner: Runner = subprocess.run,
    ) -> None:
        if not command:
            raise OCRConfigurationError("command adapter requires a non-empty argv sequence")
        if not any("{image}" in item for item in command):
            raise OCRConfigurationError("command must contain an {image} placeholder")
        self.command = tuple(str(item) for item in command)
        self.name = backend_name
        self.cwd = Path(cwd).resolve() if cwd else None
        self.env = dict(env or {})
        self.scrub_env_vars = tuple(dict.fromkeys(name for name in scrub_env_vars if name))
        self.timeout = timeout
        self.retry = retry or RetryPolicy()
        self._runner = runner

    def recognize(self, request: OCRRequest) -> OCRResult:
        with _materialize_image(request) as materialized:

            def attempt() -> OCRResult:
                with tempfile.TemporaryDirectory(prefix="pptx-wiki-command-") as output_name:
                    output_path = Path(output_name) / "result.json"
                    substitutions = {
                        "image": str(materialized.path),
                        "output": str(output_path),
                        "task": request.task,
                        "language": request.language_hint,
                        "request_id": request.request_id or "",
                    }
                    argv = [item.format_map(substitutions) for item in self.command]
                    child_env = os.environ.copy()
                    for name in self.scrub_env_vars:
                        child_env.pop(name, None)
                    child_env.update(self.env)
                    try:
                        completed = self._runner(
                            argv,
                            cwd=str(self.cwd) if self.cwd else None,
                            env=child_env,
                            capture_output=True,
                            text=True,
                            timeout=self.timeout,
                            check=False,
                        )
                    except FileNotFoundError as exc:
                        raise OCRConfigurationError(f"command executable was not found: {argv[0]!r}") from exc
                    except subprocess.TimeoutExpired as exc:
                        raise OCRTransientError(f"OCR command timed out after {self.timeout:g}s") from exc
                    except OSError as exc:
                        raise OCRTransientError(f"unable to launch OCR command: {exc}") from exc
                    if completed.returncode != 0:
                        detail = (completed.stderr or completed.stdout or "").strip()[-2000:]
                        raise OCRTransientError(f"OCR command exited with {completed.returncode}: {detail}")

                    raw_text = output_path.read_text(encoding="utf-8") if output_path.is_file() else completed.stdout
                    if not raw_text.strip():
                        raise OCRTransientError("OCR command produced no output")
                    try:
                        payload = _extract_json(raw_text)
                    except OCRParseError:
                        return OCRResult(
                            backend=self.name,
                            text=raw_text.strip(),
                            markdown=raw_text.strip(),
                            raw=raw_text,
                            warnings=["Command output was plain text/Markdown rather than structured JSON."],
                        )
                    return normalize_ocr_payload(payload, backend=self.name)

            return _retry(attempt, self.retry)


OpenAITransport = Callable[[str, Mapping[str, Any], Mapping[str, str], float], Mapping[str, Any]]


_VLM_JSON_SCHEMA: dict[str, Any] = {
    "name": "ocr_result",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "text": {"type": "string"},
            "markdown": {"type": "string"},
            "html": {"type": ["string", "null"]},
            "confidence": {"type": ["number", "null"]},
            "blocks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "kind": {"type": "string"},
                        "text": {"type": "string"},
                        "markdown": {"type": ["string", "null"]},
                        "html": {"type": ["string", "null"]},
                        "bbox": {
                            "anyOf": [
                                {"type": "array", "items": {"type": "number"}, "minItems": 4, "maxItems": 4},
                                {"type": "null"},
                            ]
                        },
                        "confidence": {"type": ["number", "null"]},
                        "order": {"type": ["integer", "null"]},
                    },
                    "required": ["kind", "text", "markdown", "html", "bbox", "confidence", "order"],
                    "additionalProperties": False,
                },
            },
        },
        "required": ["text", "markdown", "html", "confidence", "blocks"],
        "additionalProperties": False,
    },
}


class OpenAICompatibleVLMAdapter:
    """Fallback OCR through a user-provided OpenAI-compatible VLM endpoint."""

    name = "openai_vlm"

    def __init__(
        self,
        *,
        base_url: str,
        model: str,
        api_key: str | None = None,
        timeout: float = 120.0,
        max_tokens: int = 8192,
        temperature: float = 0.0,
        response_format: Literal["none", "json_object", "json_schema"] = "none",
        image_detail: Literal["auto", "low", "high"] = "high",
        retry: RetryPolicy | None = None,
        extra_body: Mapping[str, Any] | None = None,
        headers: Mapping[str, str] | None = None,
        transport: OpenAITransport | None = None,
    ) -> None:
        if not base_url.startswith(("http://", "https://")):
            raise OCRConfigurationError("base_url must be an HTTP(S) URL")
        if not model:
            raise OCRConfigurationError("model is required")
        endpoint = base_url.rstrip("/")
        if endpoint.endswith("/chat/completions"):
            self.url = endpoint
        elif endpoint.endswith("/v1"):
            self.url = endpoint + "/chat/completions"
        else:
            self.url = endpoint + "/v1/chat/completions"
        self.model = model
        self.api_key = api_key
        self.timeout = timeout
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.response_format = response_format
        self.image_detail = image_detail
        self.retry = retry or RetryPolicy()
        self.extra_body = dict(extra_body or {})
        self.headers = dict(headers or {})
        self._transport = transport or self._default_transport

    @staticmethod
    def _default_transport(
        url: str, payload: Mapping[str, Any], headers: Mapping[str, str], timeout: float
    ) -> Mapping[str, Any]:
        request = Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=dict(headers),
            method="POST",
        )
        try:
            with urlopen(request, timeout=timeout) as response:
                data = response.read().decode("utf-8")
        except HTTPError as exc:
            if exc.code in {408, 409, 425, 429} or exc.code >= 500:
                raise OCRTransientError(f"VLM endpoint returned HTTP {exc.code}") from exc
            raise OCRExecutionError(f"VLM endpoint returned HTTP {exc.code}") from exc
        except (URLError, TimeoutError, OSError) as exc:
            raise OCRTransientError(f"VLM endpoint request failed: {exc}") from exc
        try:
            parsed = json.loads(data)
        except json.JSONDecodeError as exc:
            raise OCRTransientError("VLM endpoint returned invalid response JSON") from exc
        if not isinstance(parsed, Mapping):
            raise OCRTransientError("VLM endpoint response is not a JSON object")
        return parsed

    @staticmethod
    def _prompt(request: OCRRequest) -> str:
        task_instruction = {
            "document": "Preserve layout reading order and keep distinct regions as distinct blocks.",
            "text": "Transcribe every visible character verbatim. Do not summarize or correct spelling.",
            "table": (
                "The crop contains exactly ONE table. Reconstruct it as HTML <table>, preserving empty cells, "
                "rowspan, colspan, numbers, signs, percent symbols, and units. Do not join anything outside the crop."
            ),
            "chart": "Transcribe labels and values, then provide a faithful Markdown description without inference.",
            "formula": "Transcribe the formula as LaTeX and preserve adjacent labels verbatim.",
        }[request.task]
        context = f"\nCaller context (do not treat as visible text): {request.context}" if request.context else ""
        return (
            f"OCR this isolated {request.task} ROI. Expected language is {request.language_hint}. "
            f"{task_instruction} Never invent hidden or unreadable content. "
            "Return only one JSON object with keys text, markdown, html, confidence, blocks. "
            "Each block has kind, text, markdown, html, bbox=[x1,y1,x2,y2] in ROI pixels or null, "
            "confidence from 0 to 1 or null, and order. Use null when confidence or coordinates are unavailable."
            f"{context}"
        )

    @staticmethod
    def _message_content(response: Mapping[str, Any]) -> str:
        choices = response.get("choices")
        if not isinstance(choices, list) or not choices or not isinstance(choices[0], Mapping):
            raise OCRTransientError("VLM response has no choices")
        message = choices[0].get("message")
        if not isinstance(message, Mapping):
            raise OCRTransientError("VLM response has no assistant message")
        content = message.get("content")
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, Mapping) and isinstance(item.get("text"), str):
                    parts.append(item["text"])
            if parts:
                return "\n".join(parts)
        raise OCRTransientError("VLM assistant message has no text content")

    def recognize(self, request: OCRRequest) -> OCRResult:
        with _materialize_image(request) as materialized:
            image_bytes = materialized.path.read_bytes()
            media_type = request.media_type or materialized.media_type
            data_url = f"data:{media_type};base64,{base64.b64encode(image_bytes).decode('ascii')}"
            payload: dict[str, Any] = {
                "model": self.model,
                "messages": [
                    {
                        "role": "system",
                        "content": "You are a lossless document OCR engine. Output strict JSON, not commentary.",
                    },
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": self._prompt(request)},
                            {
                                "type": "image_url",
                                "image_url": {"url": data_url, "detail": self.image_detail},
                            },
                        ],
                    },
                ],
                "temperature": self.temperature,
                "max_tokens": self.max_tokens,
            }
            if self.response_format == "json_object":
                payload["response_format"] = {"type": "json_object"}
            elif self.response_format == "json_schema":
                payload["response_format"] = {"type": "json_schema", "json_schema": _VLM_JSON_SCHEMA}
            payload.update(self.extra_body)
            headers = {"Content-Type": "application/json", **self.headers}
            if self.api_key:
                headers["Authorization"] = f"Bearer {self.api_key}"

            def attempt() -> OCRResult:
                response = self._transport(self.url, payload, headers, self.timeout)
                content = self._message_content(response)
                try:
                    parsed = _extract_json(content)
                    result = normalize_ocr_payload(parsed, backend=self.name)
                    result.raw = {"response": response, "parsed": parsed}
                    if result.confidence is not None:
                        result.warnings.append(
                            "VLM confidence is self-reported and must not be treated as a calibrated probability."
                        )
                    return result
                except OCRParseError as exc:
                    raise OCRTransientError(f"VLM returned an invalid OCR schema: {exc}") from exc

            return _retry(attempt, self.retry)


_WORKER_PROTOCOL = "pptx-wiki-ocr-worker/1"
_WORKER_FRAME_PREFIX = "@@PPTX_WIKI@@"


class PersistentOCRWorkerAdapter:
    """Run a bundled model in its own long-lived Python environment.

    The worker is intentionally a subprocess: each supported OCR model has a
    different, tightly pinned dependency set.  A newline-delimited JSON protocol
    keeps the model loaded between ROI requests while preventing those heavy
    dependencies from entering the main ``pptx-wiki`` environment.
    """

    name = "local_model"

    def __init__(
        self,
        command: Sequence[str],
        *,
        backend_name: str,
        startup_timeout: float = 900.0,
        request_timeout: float = 600.0,
        retry: RetryPolicy | None = None,
        cwd: str | Path | None = None,
        scrub_env_vars: Sequence[str] = (),
        popen_factory: Callable[..., subprocess.Popen[str]] = subprocess.Popen,
    ) -> None:
        if not command or not all(isinstance(token, str) and token for token in command):
            raise OCRConfigurationError("persistent worker command must contain non-empty strings")
        if startup_timeout <= 0 or request_timeout <= 0:
            raise OCRConfigurationError("worker timeouts must be positive")
        self.command = tuple(command)
        self.name = backend_name
        self.startup_timeout = float(startup_timeout)
        self.request_timeout = float(request_timeout)
        self.retry = retry or RetryPolicy(attempts=2, initial_delay=0.5)
        self.cwd = Path(cwd).resolve() if cwd is not None else None
        self.scrub_env_vars = tuple(dict.fromkeys(name for name in scrub_env_vars if name))
        self._popen_factory = popen_factory
        self._process: subprocess.Popen[str] | None = None
        self._stdout_queue: queue.Queue[str | None] = queue.Queue()
        self._stderr_tail: deque[str] = deque(maxlen=80)
        self._lock = threading.RLock()
        self._ready: Mapping[str, Any] | None = None

    @property
    def ready_metadata(self) -> Mapping[str, Any] | None:
        return self._ready

    def _child_environment(self) -> dict[str, str]:
        environment = os.environ.copy()
        for name in (
            *self.scrub_env_vars,
            "OPENAI_API_KEY",
            "HF_TOKEN",
            "HUGGING_FACE_HUB_TOKEN",
        ):
            environment.pop(name, None)
        environment.update(
            {
                "PYTHONUNBUFFERED": "1",
                "PYTHONUTF8": "1",
                "HF_HUB_OFFLINE": "1",
                "HF_HUB_DISABLE_IMPLICIT_TOKEN": "1",
                "HF_HUB_DISABLE_TELEMETRY": "1",
                "TRANSFORMERS_OFFLINE": "1",
                "HF_DATASETS_OFFLINE": "1",
                "TOKENIZERS_PARALLELISM": "false",
                "PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK": "True",
            }
        )
        return environment

    @staticmethod
    def _read_stream(stream: Any, target: queue.Queue[str | None]) -> None:
        try:
            for line in iter(stream.readline, ""):
                target.put(line.rstrip("\r\n"))
        finally:
            target.put(None)

    def _read_stderr(self, stream: Any) -> None:
        try:
            for line in iter(stream.readline, ""):
                value = line.rstrip("\r\n")
                if value:
                    self._stderr_tail.append(value)
        finally:
            try:
                stream.close()
            except OSError:
                pass

    def _start(self) -> None:
        if self._process is not None and self._process.poll() is None and self._ready is not None:
            return
        self.close()
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
        try:
            process = self._popen_factory(
                list(self.command),
                cwd=str(self.cwd) if self.cwd is not None else None,
                env=self._child_environment(),
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
        except FileNotFoundError as exc:
            raise OCRConfigurationError(f"local OCR worker executable was not found: {self.command[0]}") from exc
        except OSError as exc:
            raise OCRConfigurationError(f"unable to start local OCR worker: {exc}") from exc
        if process.stdin is None or process.stdout is None or process.stderr is None:
            process.kill()
            raise OCRConfigurationError("local OCR worker pipes could not be created")
        self._process = process
        self._stdout_queue = queue.Queue()
        self._stderr_tail.clear()
        self._ready = None
        threading.Thread(
            target=self._read_stream,
            args=(process.stdout, self._stdout_queue),
            name=f"{self.name}-stdout",
            daemon=True,
        ).start()
        threading.Thread(
            target=self._read_stderr,
            args=(process.stderr,),
            name=f"{self.name}-stderr",
            daemon=True,
        ).start()
        deadline = time.monotonic() + self.startup_timeout
        while True:
            message = self._next_protocol_message(deadline, phase="startup")
            if message.get("type") in {"fatal", "error"} or message.get("ok") is False:
                error = message.get("error")
                detail = error.get("message") if isinstance(error, Mapping) else error
                self.close()
                raise OCRConfigurationError(
                    f"{self.name} worker initialization failed: {detail or 'unspecified error'}"
                )
            if message.get("type") != "ready":
                continue
            self._ready = message
            return

    @staticmethod
    def _decode_protocol_line(line: str) -> Mapping[str, Any] | None:
        candidate = line
        if _WORKER_FRAME_PREFIX in line:
            candidate = line.split(_WORKER_FRAME_PREFIX, 1)[1]
        elif not line.lstrip().startswith("{"):
            return None
        try:
            value = json.loads(candidate)
        except json.JSONDecodeError:
            return None
        if not isinstance(value, Mapping) or value.get("protocol") != _WORKER_PROTOCOL:
            return None
        return value

    def _diagnostic_tail(self) -> str:
        detail = " | ".join(self._stderr_tail)
        return detail[-4000:] if detail else "no worker diagnostics"

    def _next_protocol_message(self, deadline: float, *, phase: str) -> Mapping[str, Any]:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise OCRTransientError(
                    f"{self.name} worker {phase} timed out; {self._diagnostic_tail()}"
                )
            try:
                line = self._stdout_queue.get(timeout=min(remaining, 0.25))
            except queue.Empty:
                process = self._process
                if process is not None and process.poll() is not None:
                    raise OCRTransientError(
                        f"{self.name} worker exited with {process.returncode} during {phase}; "
                        f"{self._diagnostic_tail()}"
                    )
                continue
            if line is None:
                process = self._process
                code = process.poll() if process is not None else None
                raise OCRTransientError(
                    f"{self.name} worker closed stdout during {phase} (exit={code}); "
                    f"{self._diagnostic_tail()}"
                )
            message = self._decode_protocol_line(line)
            if message is not None:
                return message

    def _request_once(self, request: OCRRequest, image_path: Path) -> OCRResult:
        self._start()
        process = self._process
        if process is None or process.stdin is None:
            raise OCRTransientError(f"{self.name} worker is not running")
        request_id = request.request_id or uuid.uuid4().hex
        payload = {
            "protocol": _WORKER_PROTOCOL,
            "type": "request",
            "op": "recognize",
            "id": request_id,
            "image": str(image_path.resolve()),
            "image_path": str(image_path.resolve()),
            "task": request.task,
            "language": request.language_hint,
            "context": request.context,
        }
        try:
            process.stdin.write(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
            process.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            self.close()
            raise OCRTransientError(f"{self.name} worker input pipe failed: {exc}") from exc

        deadline = time.monotonic() + self.request_timeout
        while True:
            message = self._next_protocol_message(deadline, phase=f"request {request_id}")
            if message.get("type") != "result":
                continue
            if message.get("id") != request_id:
                self.close()
                raise OCRExecutionError(
                    f"{self.name} worker returned an unexpected request id: "
                    f"{message.get('id')!r} (expected {request_id!r})"
                )
            if message.get("ok") is not True:
                error = message.get("error")
                if isinstance(error, Mapping):
                    error_type = str(error.get("code") or error.get("type") or "WorkerError")
                    error_message = str(error.get("message") or "unspecified worker error")
                    retryable = error.get("retryable") is True
                else:
                    error_type, error_message = "WorkerError", str(error or "unspecified worker error")
                    retryable = False
                failure = f"{self.name} {error_type}: {error_message}"
                if retryable:
                    raise OCRTransientError(failure)
                raise OCRExecutionError(failure)
            result_payload = message.get("result")
            result = normalize_ocr_payload(result_payload, backend=self.name)
            result.raw = message
            metadata = dict(self._ready or {})
            metadata.pop("protocol", None)
            metadata.pop("type", None)
            if metadata:
                result.warnings.append(
                    "Local model provenance: "
                    + ", ".join(f"{key}={value}" for key, value in sorted(metadata.items()))
                )
            return result

    def recognize(self, request: OCRRequest) -> OCRResult:
        with self._lock, _materialize_image(request) as materialized:
            def attempt() -> OCRResult:
                try:
                    return self._request_once(request, materialized.path)
                except OCRTransientError:
                    self.close()
                    raise

            return _retry(attempt, self.retry)

    def close(self) -> None:
        with self._lock:
            process = self._process
            self._process = None
            self._ready = None
            if process is None:
                return
            if process.poll() is None and process.stdin is not None:
                try:
                    process.stdin.write(
                        json.dumps(
                            {"protocol": _WORKER_PROTOCOL, "type": "shutdown", "op": "shutdown"},
                            separators=(",", ":"),
                        )
                        + "\n"
                    )
                    process.stdin.flush()
                    process.stdin.close()
                except (BrokenPipeError, OSError):
                    pass
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.terminate()
                try:
                    process.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=3)


class FallbackOCRAdapter:
    """Try adapters in order when a backend fails or returns an empty result."""

    name = "fallback"

    def __init__(self, adapters: Sequence[OCRAdapter]) -> None:
        if not adapters:
            raise OCRConfigurationError("fallback adapter requires at least one backend")
        self.adapters = tuple(adapters)

    def recognize(self, request: OCRRequest) -> OCRResult:
        failures: list[str] = []
        for adapter in self.adapters:
            try:
                result = adapter.recognize(request)
            except OCRError as exc:
                failures.append(f"{adapter.name}: {exc}")
                continue
            if result.is_empty:
                failures.append(f"{adapter.name}: empty result")
                continue
            if failures:
                result.warnings.extend(f"Fallback after {failure}" for failure in failures)
            return result
        raise OCRExecutionError("all OCR backends failed: " + "; ".join(failures))

    def close(self) -> None:
        for adapter in self.adapters:
            close = getattr(adapter, "close", None)
            if callable(close):
                close()


def create_ocr_adapter(kind: str, **config: Any) -> OCRAdapter:
    """Create one of the supported adapters.

    Stable factory names are ``paddle_cli``, ``command``, and ``openai_vlm``.
    """

    normalized = kind.strip().lower().replace("-", "_")
    factories: dict[str, Callable[..., OCRAdapter]] = {
        "paddle_cli": PaddleOCRCLIAdapter,
        "command": CommandOCRAdapter,
        "openai_vlm": OpenAICompatibleVLMAdapter,
        "persistent_worker": PersistentOCRWorkerAdapter,
    }
    try:
        factory = factories[normalized]
    except KeyError as exc:
        raise OCRConfigurationError(
            f"unknown OCR adapter {kind!r}; choose one of: {', '.join(sorted(factories))}"
        ) from exc
    return factory(**config)


__all__ = [
    "CommandOCRAdapter",
    "FallbackOCRAdapter",
    "OCRAdapter",
    "OCRBlock",
    "OCRConfigurationError",
    "OCRError",
    "OCRExecutionError",
    "OCRParseError",
    "OCRRequest",
    "OCRResult",
    "OCRTransientError",
    "OpenAICompatibleVLMAdapter",
    "PaddleOCRCLIAdapter",
    "PersistentOCRWorkerAdapter",
    "RetryPolicy",
    "create_ocr_adapter",
    "normalize_ocr_payload",
]
