"""Deterministically publish validated semantic documents as a Markdown wiki.

The publisher is deliberately model-free.  Its inputs are a semantic artifact
and the parsed provenance corpus that artifact was derived from.  Both inputs
are content-addressed, and every semantic document is validated completely
before the output directory is created or modified.
"""

from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
import json
import os
from pathlib import Path
import re
from typing import Any, Mapping, Sequence

from .semantic import SEMANTIC_SCHEMA_VERSION, load_semantic_documents
from .wiki_output import load_provenance


WIKI_SCHEMA_VERSION = "pptx-wiki.wiki.v1"
_DIGEST_RE = re.compile(r"^[0-9a-fA-F]{64}$")


@dataclass(frozen=True, slots=True)
class WikiExport:
    output_dir: Path
    index_path: Path
    report_path: Path
    page_paths: tuple[Path, ...]
    page_count: int


@dataclass(frozen=True, slots=True)
class _SemanticDocument:
    document_id: str
    title: str
    body_markdown: str
    citations: tuple[str, ...]
    content_sha256: str


def publish_wiki(
    semantic_dir: str | Path,
    parsed_corpus_dir: str | Path,
    output_dir: str | Path,
) -> WikiExport:
    """Publish a semantic artifact without invoking an LLM.

    All semantic/provenance hashes, document hashes, document identifiers and
    citations are checked before anything is written.  ``output_dir`` may be
    absent or an existing empty directory; a file or non-empty directory is
    rejected to avoid overwriting an earlier publication.
    """

    semantic_root = Path(semantic_dir).expanduser().resolve()
    parsed_root = Path(parsed_corpus_dir).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()

    manifest_path = semantic_root / "manifest.json"
    provenance_path = parsed_root / "provenance.jsonl"
    manifest = _load_manifest(manifest_path)
    validated_documents = load_semantic_documents(semantic_root)
    documents_name = str(manifest["documents_file"])
    documents_path = semantic_root / documents_name

    provenance_bytes = _read_bytes(provenance_path, "parsed provenance")
    provenance_digest = sha256(provenance_bytes).hexdigest()
    expected_provenance_digest = _required_digest(
        manifest, "source_provenance_sha256"
    )
    if provenance_digest != expected_provenance_digest:
        raise ValueError(
            "parsed provenance SHA-256 mismatch: "
            f"expected {expected_provenance_digest}, got {provenance_digest}"
        )

    documents_bytes = _read_bytes(documents_path, "semantic documents")
    documents_digest = sha256(documents_bytes).hexdigest()
    expected_documents_digest = _required_digest(manifest, "documents_sha256")
    if documents_digest != expected_documents_digest:
        raise ValueError(
            "semantic documents SHA-256 mismatch: "
            f"expected {expected_documents_digest}, got {documents_digest}"
        )

    provenance_records = load_provenance(provenance_path)
    records_by_citation = _index_provenance(provenance_records)
    documents = _documents_from_validated(validated_documents)
    _validate_documents(documents, records_by_citation, parsed_root)

    # Render everything in memory before checking/creating the output.  A
    # validation failure therefore never leaves a partially published wiki.
    rendered_pages = [
        (
            document,
            _render_page(document, records_by_citation, parsed_root, destination),
        )
        for document in documents
    ]
    index_text = _render_index(documents)

    _ensure_empty_output(destination)
    destination.mkdir(parents=True, exist_ok=True)

    page_paths: list[Path] = []
    page_records: list[dict[str, Any]] = []
    for document, page_text in rendered_pages:
        page_path = destination / f"{document.document_id}.md"
        _write_text(page_path, page_text)
        page_paths.append(page_path)
        page_records.append(
            {
                "id": document.document_id,
                "title": document.title,
                "file": page_path.name,
                "citations": list(document.citations),
                "sha256": sha256(page_text.encode("utf-8")).hexdigest(),
            }
        )

    index_path = destination / "index.md"
    _write_text(index_path, index_text)
    report = {
        "schema_version": WIKI_SCHEMA_VERSION,
        "source_semantic_schema_version": SEMANTIC_SCHEMA_VERSION,
        "source_semantic_manifest_sha256": sha256(
            manifest_path.read_bytes()
        ).hexdigest(),
        "source_provenance_sha256": provenance_digest,
        "source_documents_sha256": documents_digest,
        "page_count": len(page_paths),
        "index_sha256": sha256(index_text.encode("utf-8")).hexdigest(),
        "pages": page_records,
    }
    report_path = destination / "publish-report.json"
    _write_text(
        report_path,
        json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
    )
    return WikiExport(
        output_dir=destination,
        index_path=index_path,
        report_path=report_path,
        page_paths=tuple(page_paths),
        page_count=len(page_paths),
    )


def _load_manifest(path: Path) -> Mapping[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FileNotFoundError(f"semantic manifest not found: {path}") from None
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid semantic manifest {path}: {exc}") from exc
    if not isinstance(value, Mapping):
        raise ValueError("semantic manifest must be a JSON object")
    if value.get("schema_version") != SEMANTIC_SCHEMA_VERSION:
        raise ValueError(
            "unsupported semantic schema: "
            f"{value.get('schema_version')!r}; expected {SEMANTIC_SCHEMA_VERSION!r}"
        )
    return value


def _read_bytes(path: Path, label: str) -> bytes:
    try:
        return path.read_bytes()
    except FileNotFoundError:
        raise FileNotFoundError(f"{label} not found: {path}") from None
    except OSError as exc:
        raise ValueError(f"unable to read {label} {path}: {exc}") from exc


def _required_digest(value: Mapping[str, Any], name: str) -> str:
    digest = value.get(name)
    if not isinstance(digest, str) or not _DIGEST_RE.fullmatch(digest):
        raise ValueError(f"semantic manifest {name} must be a 64-character SHA-256 hex digest")
    return digest.casefold()


def _index_provenance(
    records: Sequence[Mapping[str, Any]],
) -> dict[str, Mapping[str, Any]]:
    indexed: dict[str, Mapping[str, Any]] = {}
    for record in records:
        citation = record.get("citation")
        if not isinstance(citation, str) or not citation:
            raise ValueError("parsed provenance contains an invalid citation")
        if citation in indexed:
            raise ValueError(f"parsed provenance contains duplicate citation: {citation}")
        indexed[citation] = record
    return indexed


def _documents_from_validated(
    values: Sequence[Mapping[str, Any]],
) -> tuple[_SemanticDocument, ...]:
    return tuple(
        _SemanticDocument(
            document_id=str(value["id"]),
            title=str(value["title"]).strip(),
            body_markdown=str(value["body_markdown"]),
            citations=tuple(dict.fromkeys(value["citations"])),
            content_sha256=str(value["content_sha256"]).casefold(),
        )
        for value in values
    )


def _validate_documents(
    documents: Sequence[_SemanticDocument],
    provenance: Mapping[str, Mapping[str, Any]],
    parsed_root: Path,
) -> None:
    for document in documents:
        unknown = [citation for citation in document.citations if citation not in provenance]
        if unknown:
            raise ValueError(
                f"semantic document {document.document_id} has unknown citation(s): "
                + ", ".join(unknown)
            )
        for citation in document.citations:
            record = provenance[citation]
            slide_number = record.get("slide_number")
            element_id = record.get("element_id")
            if (
                isinstance(slide_number, bool)
                or not isinstance(slide_number, int)
                or slide_number <= 0
                or not isinstance(element_id, str)
                or not element_id
            ):
                raise ValueError(f"parsed provenance record is invalid for citation {citation}")
            slide_path = parsed_root / "slides" / f"slide-{slide_number:04d}.md"
            if not slide_path.is_file():
                raise ValueError(
                    f"parsed slide Markdown is missing for citation {citation}: {slide_path}"
                )


def _render_page(
    document: _SemanticDocument,
    provenance: Mapping[str, Mapping[str, Any]],
    parsed_root: Path,
    output_root: Path,
) -> str:
    body = document.body_markdown.rstrip()
    lines = [f"# {document.title}", ""]
    if body:
        lines.extend((body, ""))
    lines.extend(("## 출처", ""))
    for citation in document.citations:
        record = provenance[citation]
        slide_number = int(record["slide_number"])
        element_id = str(record["element_id"])
        slide_path = parsed_root / "slides" / f"slide-{slide_number:04d}.md"
        try:
            relative = Path(os.path.relpath(slide_path, start=output_root)).as_posix()
        except ValueError as exc:
            raise ValueError(
                "wiki output and parsed corpus must be on the same filesystem "
                "to create relative source links"
            ) from exc
        lines.append(f"- {citation} — [슬라이드 원문]({relative}#{element_id})")
    return "\n".join(lines).rstrip() + "\n"


def _render_index(documents: Sequence[_SemanticDocument]) -> str:
    lines = ["# Wiki", ""]
    for document in documents:
        title = _escape_markdown_label(document.title)
        citations = ", ".join(document.citations)
        lines.append(f"- [{title}]({document.document_id}.md) — {citations}")
    return "\n".join(lines).rstrip() + "\n"


def _escape_markdown_label(value: str) -> str:
    return value.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")


def _ensure_empty_output(path: Path) -> None:
    if path.exists() and not path.is_dir():
        raise ValueError(f"wiki output exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()):
        raise ValueError(f"wiki output directory is not empty: {path}")


def _write_text(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="\n")


__all__ = ["WikiExport", "publish_wiki"]
