from __future__ import annotations

from dataclasses import asdict, dataclass, is_dataclass
from hashlib import sha256
import json
import os
from pathlib import Path
from typing import Any, Iterable

from PIL import Image

from .extract.native import extract_pptx
from .models import BBox, DeckRecord, Element, SlideRecord
from .ocr import OCRAdapter, OCRBlock, OCRError, OCRRequest, OCRResult
from .render import create_element_crops, render_pptx, render_pptx_powerpoint
from .safety import validate_pptx_archive
from .synthesis import ChatBackend, SynthesisConfig, WikiSynthesis, synthesize_wiki
from .semantic import SemanticExport
from .validate import ValidationIssue, issues_to_dicts, validate_deck
from .wiki_output import CorpusExport, export_slide_corpus


OCR_KINDS = {"image", "picture", "chart", "diagram", "media", "ole", "unknown"}


@dataclass(frozen=True, slots=True)
class PipelineConfig:
    render_backend: str = "auto"
    dpi: int = 300
    source_padding_ratio: float = 0.002
    model_padding_px: int = 24
    include_empty_shapes: bool = False
    strict_extraction: bool = False
    strict_ocr: bool = False
    office_binary: str | None = None
    pdf_binary: str | None = None
    scrub_env_vars: tuple[str, ...] = ()
    block_external_resources: bool = False


@dataclass(frozen=True, slots=True)
class PipelineResult:
    output_dir: Path
    parsed_dir: Path
    parsed_manifest_path: Path
    deck_path: Path
    corpus: CorpusExport
    qa_path: Path
    issues: tuple[ValidationIssue, ...]
    semantic: SemanticExport | None
    wiki: WikiSynthesis | None
    ocr_successes: int
    ocr_failures: int


def run_pipeline(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    config: PipelineConfig | None = None,
    ocr_adapter: OCRAdapter | None = None,
    rendered_slides_dir: str | Path | None = None,
    synthesis_backend: ChatBackend | None = None,
    synthesis_config: SynthesisConfig | None = None,
) -> PipelineResult:
    """Run parsed extraction and optionally continue through semantic and Wiki stages."""

    settings = config or PipelineConfig()
    source = Path(pptx_path).resolve()
    destination = Path(output_dir).resolve()
    parsed = destination / "parsed"
    validate_pptx_archive(
        source,
        reject_external_resources=settings.block_external_resources,
    )
    _ensure_empty_stage(parsed, "parsed")
    _ensure_empty_stage(destination / "semantic", "semantic")
    _ensure_empty_stage(destination / "wiki", "wiki")
    parsed.mkdir(parents=True, exist_ok=True)

    source_assets = parsed / "source-assets"
    deck = extract_pptx(
        source,
        assets_dir=source_assets,
        include_empty_shapes=settings.include_empty_shapes,
        strict=settings.strict_extraction,
    )

    ocr_successes = 0
    ocr_failures = 0
    if ocr_adapter is not None or rendered_slides_dir is not None:
        rendered = _resolve_rendered_slides(
            source,
            parsed,
            len(deck.slides),
            rendered_slides_dir=rendered_slides_dir,
            dpi=settings.dpi,
            backend=settings.render_backend,
            office_binary=settings.office_binary,
            pdf_binary=settings.pdf_binary,
            scrub_env_vars=settings.scrub_env_vars,
        )
        for slide, rendered_path in zip(deck.slides, rendered, strict=True):
            slide.rendered_path = str(rendered_path)
        create_element_crops(
            deck,
            rendered,
            parsed / "roi",
            element_kinds=OCR_KINDS,
            padding_ratio=settings.source_padding_ratio,
            model_padding_px=settings.model_padding_px,
        )
        if ocr_adapter is not None:
            ocr_successes, ocr_failures = apply_ocr(
                deck,
                ocr_adapter,
                output_dir=parsed / "ocr-results",
                strict=settings.strict_ocr,
            )

    deck_path = parsed / "deck.json"
    _write_json(deck_path, deck.to_dict())
    corpus = export_slide_corpus(deck, parsed / "corpus")
    issues = validate_deck(deck)
    qa_path = parsed / "qa.json"
    _write_json(
        qa_path,
        {
            "issue_count": len(issues),
            "error_count": sum(issue.severity == "error" for issue in issues),
            "ocr_successes": ocr_successes,
            "ocr_failures": ocr_failures,
            "extraction_warnings": deck.metadata.get("warnings", []),
            "issues": issues_to_dicts(issues),
        },
    )
    parsed_manifest_path = parsed / "manifest.json"
    _write_json(
        parsed_manifest_path,
        {
            "schema_version": "pptx-wiki.parsed.v1",
            "source": {
                "name": source.name,
                "size": source.stat().st_size,
                "sha256": _file_sha256(source),
            },
            "paths": {
                "deck": deck_path.relative_to(parsed).as_posix(),
                "qa": qa_path.relative_to(parsed).as_posix(),
                "corpus": corpus.output_dir.relative_to(parsed).as_posix(),
                "provenance": corpus.provenance_path.relative_to(parsed).as_posix(),
            },
            "files": {
                "deck_sha256": sha256(deck_path.read_bytes()).hexdigest(),
                "qa_sha256": sha256(qa_path.read_bytes()).hexdigest(),
                "provenance_sha256": corpus.digest,
            },
            "slide_count": corpus.slide_count,
            "block_count": corpus.block_count,
            "qa_error_count": sum(issue.severity == "error" for issue in issues),
            "qa_issue_count": len(issues),
            "ocr_successes": ocr_successes,
            "ocr_failures": ocr_failures,
        },
    )

    semantic = None
    wiki = None
    if synthesis_backend is not None:
        wiki = synthesize_wiki(
            corpus.output_dir,
            backend=synthesis_backend,
            output_dir=destination / "wiki",
            config=synthesis_config,
            semantic_output_dir=destination / "semantic",
        )
        semantic = wiki.semantic
    return PipelineResult(
        output_dir=destination,
        parsed_dir=parsed,
        parsed_manifest_path=parsed_manifest_path,
        deck_path=deck_path,
        corpus=corpus,
        qa_path=qa_path,
        issues=tuple(issues),
        semantic=semantic,
        wiki=wiki,
        ocr_successes=ocr_successes,
        ocr_failures=ocr_failures,
    )


def apply_ocr(
    deck: DeckRecord,
    adapter: OCRAdapter,
    *,
    output_dir: str | Path,
    strict: bool = False,
) -> tuple[int, int]:
    """OCR only raster/visual elements and append results as child elements."""

    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    successes = 0
    failures: list[dict[str, str]] = []

    for slide in deck.slides:
        if not slide.rendered_path:
            continue
        with Image.open(slide.rendered_path) as rendered_image:
            rendered_size = rendered_image.size
        targets = [
            element
            for element in list(slide.elements)
            if element.kind in OCR_KINDS
            and element.metadata.get("ocr_policy") != "never"
            and element.asset_path
        ]
        additions: list[Element] = []
        for element in targets:
            request = OCRRequest(
                image=element.asset_path,
                task=_task_for_element(element),
                language_hint="ko",
                context=(
                    f"PowerPoint slide {slide.number}; source element {element.id}. "
                    "Keep visually separate tables as separate blocks."
                ),
                request_id=element.id,
            )
            try:
                result = adapter.recognize(request)
            except OCRError as exc:
                failures.append({"element_id": element.id, "error": f"{type(exc).__name__}: {exc}"})
                if strict:
                    raise
                continue
            additions.extend(_ocr_elements(element, slide, result, rendered_size))
            _write_json(destination / f"{element.id}.json", result.to_dict())
            element.metadata["ocr_backend"] = result.backend
            element.metadata["ocr_child_count"] = len(result.blocks) or int(not result.is_empty)
            element.metadata["ocr_warnings"] = list(result.warnings)
            successes += 1
        slide.elements.extend(additions)

    deck.metadata["ocr"] = {
        "adapter": getattr(adapter, "name", type(adapter).__name__),
        "successes": successes,
        "failures": failures,
    }
    return successes, len(failures)


def _ocr_elements(
    parent: Element,
    slide: SlideRecord,
    result: OCRResult,
    rendered_size: tuple[int, int],
) -> list[Element]:
    blocks: Iterable[OCRBlock]
    if result.blocks:
        blocks = result.blocks
    elif result.is_empty:
        return []
    else:
        blocks = [
            OCRBlock(
                kind=_task_for_element(parent),
                text=result.text,
                markdown=result.markdown or None,
                html=result.html,
                confidence=result.confidence,
            )
        ]

    source = "vlm" if result.backend == "openai_vlm" else "ocr"
    elements: list[Element] = []
    for index, block in enumerate(blocks, start=1):
        kind_token = "".join(character if character.isalnum() else "_" for character in block.kind.casefold()).strip("_")
        kind = f"ocr_{kind_token or 'block'}"
        elements.append(
            Element(
                id=f"{parent.id}-ocr{index:03d}",
                slide_number=slide.number,
                kind=kind,
                bbox=_block_bbox_on_slide(parent, block, slide, rendered_size),
                z_index=parent.z_index + index,
                source=source,
                name=f"{parent.name or parent.id} / OCR block {index}",
                text=block.text or None,
                markdown=block.markdown,
                html=block.html,
                confidence=block.confidence,
                parent_id=parent.id,
                metadata={
                    "ocr_backend": result.backend,
                    "ocr_order": block.order,
                    "confidence_source": block.confidence_source,
                    "ocr_block_metadata": block.metadata,
                    "parent_crop": parent.asset_path,
                },
            )
        )
    return elements


def _block_bbox_on_slide(
    parent: Element,
    block: OCRBlock,
    slide: SlideRecord,
    rendered_size: tuple[int, int],
) -> BBox:
    crop = parent.metadata.get("ocr_crop_bbox_px")
    if block.bbox is None or not isinstance(crop, dict):
        return parent.bbox
    rendered_width, rendered_height = rendered_size
    padding = int(parent.metadata.get("model_padding_px", 0) or 0)
    crop_left = int(crop["left"])
    crop_top = int(crop["top"])
    crop_right = int(crop["right"])
    crop_bottom = int(crop["bottom"])
    x1, y1, x2, y2 = block.bbox
    local_left = max(0.0, min(crop_right - crop_left, x1 - padding))
    local_top = max(0.0, min(crop_bottom - crop_top, y1 - padding))
    local_right = max(local_left, min(crop_right - crop_left, x2 - padding))
    local_bottom = max(local_top, min(crop_bottom - crop_top, y2 - padding))
    px_left = crop_left + local_left
    px_top = crop_top + local_top
    px_right = crop_left + local_right
    px_bottom = crop_top + local_bottom
    emu_left = round(px_left / rendered_width * slide.width)
    emu_top = round(px_top / rendered_height * slide.height)
    emu_right = round(px_right / rendered_width * slide.width)
    emu_bottom = round(px_bottom / rendered_height * slide.height)
    if emu_right <= emu_left or emu_bottom <= emu_top:
        return parent.bbox
    return BBox(emu_left, emu_top, emu_right - emu_left, emu_bottom - emu_top)


def _task_for_element(element: Element) -> str:
    kind = element.kind.casefold()
    if "table" in kind:
        return "table"
    if "chart" in kind:
        return "chart"
    if "formula" in kind or "equation" in kind:
        return "formula"
    if kind in {"text", "textbox"}:
        return "text"
    return "document"


def _resolve_rendered_slides(
    source: Path,
    destination: Path,
    expected_count: int,
    *,
    rendered_slides_dir: str | Path | None,
    backend: str,
    dpi: int,
    office_binary: str | None,
    pdf_binary: str | None,
    scrub_env_vars: tuple[str, ...],
) -> list[Path]:
    if rendered_slides_dir is None:
        selected_backend = backend.casefold()
        if selected_backend == "auto":
            selected_backend = "powerpoint" if os.name == "nt" else "libreoffice"
        if selected_backend == "powerpoint":
            slides = render_pptx_powerpoint(
                source,
                destination / "rendered",
                dpi=dpi,
                scrub_env_vars=scrub_env_vars,
            )
        elif selected_backend == "libreoffice":
            slides = render_pptx(
                source,
                destination / "rendered",
                dpi=dpi,
                office_binary=office_binary,
                pdf_binary=pdf_binary,
                scrub_env_vars=scrub_env_vars,
            )
        else:
            raise ValueError("render backend must be auto, powerpoint, or libreoffice")
    else:
        root = Path(rendered_slides_dir)
        slides = sorted(
            (path for path in root.iterdir() if path.suffix.casefold() in {".png", ".jpg", ".jpeg", ".webp"}),
            key=_natural_slide_key,
        )
    if len(slides) != expected_count:
        raise ValueError(f"expected {expected_count} rendered slides, found {len(slides)}")
    return [path.resolve() for path in slides]


def _natural_slide_key(path: Path) -> tuple[int, str]:
    digits = "".join(character for character in path.stem if character.isdigit())
    return (int(digits) if digits else 10**12, path.name)


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2, default=_json_default) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def _ensure_empty_stage(path: Path, label: str) -> None:
    if path.exists() and not path.is_dir():
        raise ValueError(f"{label} output exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()):
        raise ValueError(f"{label} output directory is not empty: {path}")


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _json_default(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, bytes):
        return {"type": "bytes", "length": len(value)}
    if is_dataclass(value):
        return asdict(value)
    return str(value)


__all__ = ["OCR_KINDS", "PipelineConfig", "PipelineResult", "apply_ocr", "run_pipeline"]
