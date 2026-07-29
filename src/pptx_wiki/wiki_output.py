"""Deterministic, evidence-preserving output for an extracted PPTX deck.

This module deliberately does not try to make the slide contents read like a
single document.  In particular, two nearby tables remain two independent
blocks.  Semantic reorganisation belongs in :mod:`pptx_wiki.synthesis`, where
every generated statement must retain a citation back to one of these blocks.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, is_dataclass
from hashlib import sha256
from html import escape as html_escape
import json
from pathlib import Path
import re
from typing import Any, Mapping, Sequence
import unicodedata

from .models import BBox, DeckRecord, Element, SlideRecord, TableData


SCHEMA_VERSION = "pptx-wiki.provenance.v1"


@dataclass(frozen=True, slots=True)
class CorpusExport:
    """Paths and stable counts produced by :func:`export_slide_corpus`."""

    output_dir: Path
    slides_dir: Path
    provenance_path: Path
    manifest_path: Path
    slide_paths: tuple[Path, ...]
    slide_count: int
    block_count: int
    digest: str


def export_slide_corpus(
    deck: DeckRecord | Mapping[str, Any],
    output_dir: str | Path,
    *,
    document_id: str | None = None,
) -> CorpusExport:
    """Write deterministic slide Markdown, provenance JSONL and a manifest.

    ``deck`` normally is a :class:`~pptx_wiki.models.DeckRecord`.  Mapping
    input is accepted as a convenience for fixtures and for pipelines that
    persist the extraction model as JSON before this stage.

    The output is intentionally loss-aware:

    * each PPTX element is one provenance record;
    * every Markdown block has an exact ``[slide-N#element-id]`` citation;
    * table elements have explicit begin/end boundaries and are never merged;
    * merged cells are rendered as HTML when GFM cannot represent their spans;
    * no timestamp is written, so identical input produces byte-identical
      output.
    """

    target = Path(output_dir)
    slides_dir = target / "slides"
    slides_dir.mkdir(parents=True, exist_ok=True)

    source_path = str(_get(deck, "source_path", ""))
    resolved_document_id = document_id or _default_document_id(source_path)
    slide_values = list(_get(deck, "slides", ()) or ())
    slide_values.sort(key=lambda slide: int(_get(slide, "number", 0)))

    all_records: list[dict[str, Any]] = []
    slide_paths: list[Path] = []

    for slide in slide_values:
        slide_number = int(_get(slide, "number", 0))
        if slide_number <= 0:
            raise ValueError(f"slide numbers must be positive, got {slide_number!r}")
        markdown, records = _render_slide(
            slide,
            document_id=resolved_document_id,
            source_path=source_path,
        )
        slide_path = slides_dir / f"slide-{slide_number:04d}.md"
        _write_text(slide_path, markdown)
        slide_paths.append(slide_path)
        all_records.extend(records)

    provenance_path = target / "provenance.jsonl"
    jsonl = "".join(_json_dumps(record) + "\n" for record in all_records)
    _write_text(provenance_path, jsonl)

    digest = sha256(jsonl.encode("utf-8")).hexdigest()
    manifest = {
        "schema_version": SCHEMA_VERSION,
        "document_id": resolved_document_id,
        "source_path": source_path,
        "slide_count": len(slide_paths),
        "block_count": len(all_records),
        "provenance_file": provenance_path.name,
        "slides": [path.relative_to(target).as_posix() for path in slide_paths],
        "provenance_sha256": digest,
    }
    manifest_path = target / "manifest.json"
    _write_text(manifest_path, _json_dumps(manifest, indent=2) + "\n")

    return CorpusExport(
        output_dir=target,
        slides_dir=slides_dir,
        provenance_path=provenance_path,
        manifest_path=manifest_path,
        slide_paths=tuple(slide_paths),
        slide_count=len(slide_paths),
        block_count=len(all_records),
        digest=digest,
    )


def load_provenance(path: str | Path) -> list[dict[str, Any]]:
    """Load and minimally validate a provenance JSONL file."""

    records: list[dict[str, Any]] = []
    with Path(path).open("r", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, start=1):
            if not line.strip():
                continue
            value = json.loads(line)
            if not isinstance(value, dict):
                raise ValueError(f"provenance line {line_number} is not an object")
            for required in ("citation", "slide_number", "element_id", "content"):
                if required not in value:
                    raise ValueError(
                        f"provenance line {line_number} is missing {required!r}"
                    )
            records.append(value)
    return records


def render_table(table: TableData | Mapping[str, Any]) -> tuple[str, str]:
    """Render a structured table without discarding merged-cell semantics.

    Returns ``(content, format)``.  Plain rectangular tables become GFM
    Markdown.  A table containing row/column spans becomes HTML because pipe
    tables cannot faithfully encode those relationships.
    """

    rows = int(_get(table, "rows", 0))
    columns = int(_get(table, "columns", 0))
    if rows <= 0 or columns <= 0:
        return "", "markdown"

    cells = list(_get(table, "cells", ()) or ())
    origins: dict[tuple[int, int], Any] = {}
    has_spans = False
    for cell in cells:
        if not bool(_get(cell, "is_merge_origin", True)):
            continue
        row = int(_get(cell, "row", 0))
        column = int(_get(cell, "column", 0))
        if not (0 <= row < rows and 0 <= column < columns):
            continue
        origins[(row, column)] = cell
        if int(_get(cell, "row_span", 1)) > 1 or int(_get(cell, "column_span", 1)) > 1:
            has_spans = True

    if has_spans:
        return _table_as_html(rows, columns, origins), "html"

    matrix: list[list[str]] = []
    for row in range(rows):
        matrix.append(
            [_normalise_text(str(_get(origins.get((row, col)), "text", ""))) for col in range(columns)]
        )
    return _matrix_as_markdown(matrix), "markdown"


def _render_slide(
    slide: SlideRecord | Mapping[str, Any],
    *,
    document_id: str,
    source_path: str,
) -> tuple[str, list[dict[str, Any]]]:
    slide_number = int(_get(slide, "number", 0))
    width = int(_get(slide, "width", 0))
    height = int(_get(slide, "height", 0))
    title = _normalise_text(str(_get(slide, "title", "") or ""))

    elements = list(_get(slide, "elements", ()) or ())
    indexed = list(enumerate(elements))
    indexed.sort(key=lambda pair: _element_order_key(pair[1], pair[0]))

    used_ids: set[str] = set()
    rendered_blocks: list[str] = []
    records: list[dict[str, Any]] = []

    for reading_order, (_, element) in enumerate(indexed, start=1):
        base_id = _safe_element_id(str(_get(element, "id", "") or f"element-{reading_order}"))
        element_id = _deduplicate_id(base_id, used_ids)
        content, content_format = _element_content(element)
        # Empty shapes are layout evidence but add no wiki evidence.  Keep
        # asset-only elements, since the asset path itself is useful provenance.
        asset_path = _get(element, "asset_path", None)
        if not content and not asset_path:
            continue
        citation = f"[slide-{slide_number}#{element_id}]"
        record = _provenance_record(
            element,
            document_id=document_id,
            source_path=source_path,
            slide_number=slide_number,
            slide_width=width,
            slide_height=height,
            slide_title=title,
            element_id=element_id,
            citation=citation,
            reading_order=reading_order,
            content=content,
            content_format=content_format,
        )
        records.append(record)
        rendered_blocks.append(_markdown_block(record))

    notes = _normalise_text(str(_get(slide, "notes", "") or ""))
    if notes:
        element_id = _deduplicate_id("notes", used_ids)
        citation = f"[slide-{slide_number}#{element_id}]"
        note_record = {
            "schema_version": SCHEMA_VERSION,
            "document_id": document_id,
            "source_path": source_path,
            "slide_number": slide_number,
            "slide_title": title,
            "slide_size_emu": {"width": width, "height": height},
            "citation": citation,
            "element_id": element_id,
            "original_element_id": None,
            "kind": "speaker_notes",
            "source": "native",
            "name": "Speaker notes",
            "reading_order": max(
                (int(record["reading_order"]) for record in records), default=0
            )
            + 1,
            "z_index": None,
            "bbox_emu": None,
            "bbox_normalized": None,
            "parent_id": None,
            "asset_path": None,
            "confidence": None,
            "content_format": "text",
            "content": notes,
            "table": None,
            "metadata": {"synthetic": True},
            "content_sha256": sha256(notes.encode("utf-8")).hexdigest(),
        }
        records.append(note_record)
        rendered_blocks.append(_markdown_block(note_record))

    frontmatter = [
        "---",
        f"schema: {SCHEMA_VERSION}",
        f"document_id: {_yaml_scalar(document_id)}",
        f"slide_number: {slide_number}",
        f"title: {_yaml_scalar(title)}",
        f"source_path: {_yaml_scalar(source_path)}",
        "---",
        "",
    ]
    heading = f"# Slide {slide_number}"
    if title:
        heading += f" — {title}"
    body = frontmatter + [heading, ""]
    if rendered_blocks:
        body.extend(["\n\n".join(rendered_blocks), ""])
    else:
        body.extend(["<!-- This slide has no extractable content. -->", ""])
    return "\n".join(body), records


def _provenance_record(
    element: Element | Mapping[str, Any],
    *,
    document_id: str,
    source_path: str,
    slide_number: int,
    slide_width: int,
    slide_height: int,
    slide_title: str,
    element_id: str,
    citation: str,
    reading_order: int,
    content: str,
    content_format: str,
) -> dict[str, Any]:
    bbox = _get(element, "bbox", None)
    bbox_emu = None
    bbox_normalized = None
    if bbox is not None:
        x = int(_get(bbox, "x", 0))
        y = int(_get(bbox, "y", 0))
        width = int(_get(bbox, "width", 0))
        height = int(_get(bbox, "height", 0))
        bbox_emu = {"x": x, "y": y, "width": width, "height": height}
        if slide_width > 0 and slide_height > 0:
            bbox_normalized = [
                round(x / slide_width, 6),
                round(y / slide_height, 6),
                round((x + width) / slide_width, 6),
                round((y + height) / slide_height, 6),
            ]

    table = _get(element, "table", None)
    original_id = str(_get(element, "id", "") or "") or None
    return {
        "schema_version": SCHEMA_VERSION,
        "document_id": document_id,
        "source_path": source_path,
        "slide_number": slide_number,
        "slide_title": slide_title,
        "slide_size_emu": {"width": slide_width, "height": slide_height},
        "citation": citation,
        "element_id": element_id,
        "original_element_id": original_id,
        "kind": str(_get(element, "kind", "unknown")),
        "source": str(_get(element, "source", "derived")),
        "name": _get(element, "name", None),
        "reading_order": reading_order,
        "z_index": _get(element, "z_index", None),
        "bbox_emu": bbox_emu,
        "bbox_normalized": bbox_normalized,
        "parent_id": _get(element, "parent_id", None),
        "asset_path": _get(element, "asset_path", None),
        "confidence": _get(element, "confidence", None),
        "content_format": content_format,
        "content": content,
        "table": _jsonable(table),
        "metadata": _jsonable(_get(element, "metadata", {}) or {}),
        "content_sha256": sha256(content.encode("utf-8")).hexdigest(),
    }


def _element_content(element: Element | Mapping[str, Any]) -> tuple[str, str]:
    kind = str(_get(element, "kind", "")).casefold()
    table = _get(element, "table", None)
    markdown = _normalise_text(str(_get(element, "markdown", "") or ""))
    html = _normalise_text(str(_get(element, "html", "") or ""))
    text = _normalise_text(str(_get(element, "text", "") or ""))

    # A native TableData is the most lossless representation.  Do not replace
    # it with an OCR markdown string merely because one is also present.
    if table is not None:
        rendered, rendered_format = render_table(table)
        if rendered:
            return rendered, rendered_format
    if markdown:
        return markdown, "markdown"
    if html:
        return html, "html"
    if text:
        return text, "text"
    asset_path = str(_get(element, "asset_path", "") or "")
    if asset_path:
        label = str(_get(element, "name", "") or kind or "asset")
        return f"![{_escape_markdown(label)}]({asset_path})", "markdown"
    return "", "text"


def _markdown_block(record: Mapping[str, Any]) -> str:
    citation = str(record["citation"])
    kind = str(record.get("kind", "unknown"))
    source = str(record.get("source", "derived"))
    bbox = record.get("bbox_normalized")
    content = str(record.get("content", ""))
    element_id = str(record["element_id"])
    bbox_token = "null" if bbox is None else ",".join(f"{float(value):.6f}" for value in bbox)
    begin = (
        f"<!-- BEGIN BLOCK {citation} kind={kind} source={source} "
        f"bbox_norm={bbox_token} -->"
    )
    anchor = f'<a id="{html_escape(element_id, quote=True)}"></a>'
    end = f"<!-- END BLOCK {citation} -->"
    if kind.casefold() in {"table", "native_table", "ocr_table"}:
        begin = (
            f"<!-- BEGIN TABLE {citation} kind={kind} source={source} "
            f"bbox_norm={bbox_token} -->"
        )
        end = f"<!-- END TABLE {citation} -->"
    return "\n".join((begin, anchor, content, "", citation, end))


def _element_order_key(element: Element | Mapping[str, Any], index: int) -> tuple[Any, ...]:
    metadata = _get(element, "metadata", {}) or {}
    explicit_order = _get(metadata, "reading_order", None)
    bbox = _get(element, "bbox", None)
    y = int(_get(bbox, "y", 0)) if bbox is not None else 0
    x = int(_get(bbox, "x", 0)) if bbox is not None else 0
    z_index = int(_get(element, "z_index", index) or 0)
    if isinstance(explicit_order, (int, float)):
        return (0, float(explicit_order), y, x, z_index, index)
    return (1, y, x, z_index, index)


def _table_as_html(rows: int, columns: int, origins: Mapping[tuple[int, int], Any]) -> str:
    covered: set[tuple[int, int]] = set()
    lines = ["<table>"]
    for row in range(rows):
        lines.append("  <tr>")
        for column in range(columns):
            if (row, column) in covered:
                continue
            cell = origins.get((row, column))
            row_span = max(1, int(_get(cell, "row_span", 1))) if cell else 1
            column_span = max(1, int(_get(cell, "column_span", 1))) if cell else 1
            for covered_row in range(row, min(rows, row + row_span)):
                for covered_column in range(column, min(columns, column + column_span)):
                    if (covered_row, covered_column) != (row, column):
                        covered.add((covered_row, covered_column))
            attributes: list[str] = []
            if row_span > 1:
                attributes.append(f'rowspan="{row_span}"')
            if column_span > 1:
                attributes.append(f'colspan="{column_span}"')
            attrs = " " + " ".join(attributes) if attributes else ""
            cell_text = _normalise_text(str(_get(cell, "text", "") if cell else ""))
            cell_html = html_escape(cell_text).replace("\n", "<br>")
            lines.append(f"    <td{attrs}>{cell_html}</td>")
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)


def _matrix_as_markdown(matrix: Sequence[Sequence[str]]) -> str:
    if not matrix:
        return ""
    width = max((len(row) for row in matrix), default=0)
    if width == 0:
        return ""

    def row(values: Sequence[str]) -> str:
        padded = list(values) + [""] * (width - len(values))
        return "| " + " | ".join(_escape_table_cell(value) for value in padded) + " |"

    header = row(matrix[0])
    delimiter = "| " + " | ".join("---" for _ in range(width)) + " |"
    rest = [row(values) for values in matrix[1:]]
    return "\n".join((header, delimiter, *rest))


def _escape_table_cell(value: str) -> str:
    return value.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _escape_markdown(value: str) -> str:
    return value.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")


def _normalise_text(value: str) -> str:
    value = unicodedata.normalize("NFC", value.replace("\r\n", "\n").replace("\r", "\n"))
    return "\n".join(line.rstrip() for line in value.split("\n")).strip()


def _safe_element_id(value: str) -> str:
    value = unicodedata.normalize("NFC", value).strip()
    # ``]``, ``#`` and whitespace would make the citation grammar ambiguous.
    value = re.sub(r"[\s\]#\[]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "element"


def _deduplicate_id(base: str, used: set[str]) -> str:
    candidate = base
    suffix = 2
    while candidate in used:
        candidate = f"{base}--{suffix}"
        suffix += 1
    used.add(candidate)
    return candidate


def _default_document_id(source_path: str) -> str:
    stem = Path(source_path).stem if source_path else "deck"
    stem = _safe_element_id(stem)
    return stem or "deck"


def _yaml_scalar(value: str) -> str:
    # A JSON string is also a valid YAML scalar and avoids a YAML dependency.
    return json.dumps(value, ensure_ascii=False)


def _jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value):
        return _jsonable(asdict(value))
    if isinstance(value, Mapping):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    if isinstance(value, set):
        return sorted((_jsonable(item) for item in value), key=repr)
    return str(value)


def _get(value: Any, name: str, default: Any = None) -> Any:
    if value is None:
        return default
    if isinstance(value, Mapping):
        return value.get(name, default)
    return getattr(value, name, default)


def _json_dumps(value: Any, *, indent: int | None = None) -> str:
    return json.dumps(
        _jsonable(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":") if indent is None else None,
        indent=indent,
    )


def _write_text(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="\n")


__all__ = [
    "CorpusExport",
    "SCHEMA_VERSION",
    "export_slide_corpus",
    "load_provenance",
    "render_table",
]
