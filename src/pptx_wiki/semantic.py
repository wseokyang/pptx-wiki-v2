"""Grounded semantic reorganisation between parsed evidence and Wiki publishing.

The parsed corpus is immutable source evidence.  This stage may use an LLM to
select and reorganise that evidence, but it persists the result as a
machine-readable artifact instead of publishing Wiki pages directly.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from hashlib import sha256
import json
from pathlib import Path
import re
from typing import Any, Callable, Mapping
import unicodedata

from .synthesis import (
    CITATION_RE,
    ChatBackend,
    SynthesisConfig,
    Topic,
    _add_uncovered_topics,
    _discover_topics,
    _generate_topic_page,
    _normalise_topics,
    _slide_topics,
    _slugify,
    _unique_slug,
    _without_top_heading,
)
from .wiki_output import load_provenance


SEMANTIC_SCHEMA_VERSION = "pptx-wiki.semantic.v1"
SemanticConfig = SynthesisConfig
_DIGEST_RE = re.compile(r"^[0-9a-fA-F]{64}$")
_DOCUMENT_ID_RE = re.compile(r"^[\w.-]+$", re.UNICODE)
_WINDOWS_RESERVED_IDS = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{number}" for number in range(1, 10)),
    *(f"lpt{number}" for number in range(1, 10)),
}


@dataclass(frozen=True, slots=True)
class SemanticDocument:
    id: str
    title: str
    description: str
    body_markdown: str
    citations: tuple[str, ...]
    generation: str
    content_sha256: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "title": self.title,
            "description": self.description,
            "body_markdown": self.body_markdown,
            "citations": list(self.citations),
            "generation": self.generation,
            "content_sha256": self.content_sha256,
        }


@dataclass(frozen=True, slots=True)
class SemanticExport:
    output_dir: Path
    manifest_path: Path
    documents_path: Path
    documents: tuple[SemanticDocument, ...]
    topics: tuple[Topic, ...]
    document_count: int
    selected_citations: tuple[str, ...]
    omitted_citations: tuple[str, ...]
    fallback_documents: tuple[str, ...]
    warnings: tuple[str, ...]


def build_semantic_output(
    corpus_dir: str | Path,
    *,
    backend: ChatBackend | Callable[..., str],
    output_dir: str | Path,
    config: SemanticConfig | None = None,
) -> SemanticExport:
    """Create a grounded semantic artifact from an immutable parsed corpus.

    ``coverage_policy='selected'`` permits the model to omit irrelevant source
    blocks.  Omitted citations remain auditable in the manifest.  ``complete``
    retains every citation and is the compatibility/fail-open policy.
    """

    corpus = Path(corpus_dir).resolve()
    provenance_path = corpus / "provenance.jsonl"
    if not provenance_path.is_file():
        raise FileNotFoundError(provenance_path)
    destination = Path(output_dir).resolve()
    if destination.exists() and any(destination.iterdir()):
        raise ValueError(f"semantic output directory is not empty: {destination}")
    destination.mkdir(parents=True, exist_ok=True)

    settings = config or SemanticConfig()
    records = load_provenance(provenance_path)
    if not records:
        raise ValueError("the parsed provenance corpus contains no evidence records")
    citations = tuple(str(record["citation"]) for record in records)
    if len(set(citations)) != len(citations):
        raise ValueError("parsed provenance contains duplicate citations")
    citation_rank = {citation: rank for rank, citation in enumerate(citations)}
    record_by_citation = {
        str(record["citation"]): record for record in records
    }

    warnings: list[str] = []
    if settings.discover_topics:
        topics = _discover_topics(records, backend, settings, warnings)
    else:
        topics = _slide_topics(records)
    topics = _normalise_topics(topics, citation_rank, settings.max_topics)
    if settings.coverage_policy == "complete":
        topics = _add_uncovered_topics(topics, records, citation_rank)
        topics = _normalise_topics(topics, citation_rank, settings.max_topics)

    selected = {
        citation
        for topic in topics
        for citation in topic.citations
        if citation in record_by_citation
    }
    selected_citations = tuple(
        citation for citation in citations if citation in selected
    )
    omitted_citations = tuple(
        citation for citation in citations if citation not in selected
    )
    if not selected_citations:
        warnings.append("semantic selection was empty; all evidence was retained")
        topics = _normalise_topics(
            _slide_topics(records), citation_rank, settings.max_topics
        )
        selected_citations = citations
        omitted_citations = ()

    documents: list[SemanticDocument] = []
    fallback_documents: list[str] = []
    # Reserve publisher-owned filenames so a semantic title can never replace
    # the Wiki index or audit report.
    used_ids: set[str] = {
        "index",
        "publish-report",
        "manifest",
        "documents",
        *_WINDOWS_RESERVED_IDS,
    }
    for topic in topics:
        topic_citations = tuple(
            citation
            for citation in topic.citations
            if citation in record_by_citation
        )
        if not topic_citations:
            continue
        topic_records = [record_by_citation[citation] for citation in topic_citations]
        page, used_fallback, page_warnings = _generate_topic_page(
            Topic(topic.title, topic_citations, topic.description),
            topic_records,
            backend,
            settings,
        )
        warnings.extend(page_warnings)
        body = _without_top_heading(page).strip()
        document_id = _unique_slug(_slugify(topic.title), used_ids)
        content_digest = sha256(body.encode("utf-8")).hexdigest()
        document = SemanticDocument(
            id=document_id,
            title=topic.title,
            description=topic.description,
            body_markdown=body,
            citations=topic_citations,
            generation="verbatim_fallback" if used_fallback else "model",
            content_sha256=content_digest,
        )
        documents.append(document)
        if used_fallback:
            fallback_documents.append(document_id)

    if not documents:
        raise ValueError("semantic stage produced no documents")

    documents_path = destination / "documents.jsonl"
    documents_text = "".join(
        json.dumps(document.to_dict(), ensure_ascii=False, sort_keys=True) + "\n"
        for document in documents
    )
    documents_path.write_text(
        documents_text, encoding="utf-8", newline="\n"
    )
    provenance_digest = sha256(provenance_path.read_bytes()).hexdigest()
    documents_digest = sha256(documents_text.encode("utf-8")).hexdigest()
    manifest = {
        "schema_version": SEMANTIC_SCHEMA_VERSION,
        "source_provenance_sha256": provenance_digest,
        "documents_file": documents_path.name,
        "documents_sha256": documents_digest,
        "document_count": len(documents),
        "record_count": len(records),
        "selected_citations": list(selected_citations),
        "omitted_citations": list(omitted_citations),
        "fallback_documents": fallback_documents,
        "backend": {
            "type": type(backend).__name__,
            "model": getattr(backend, "model", None),
        },
        "config": asdict(settings),
        "warnings": warnings,
    }
    manifest_path = destination / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return SemanticExport(
        output_dir=destination,
        manifest_path=manifest_path,
        documents_path=documents_path,
        documents=tuple(documents),
        topics=tuple(
            Topic(document.title, document.citations, document.description)
            for document in documents
        ),
        document_count=len(documents),
        selected_citations=selected_citations,
        omitted_citations=omitted_citations,
        fallback_documents=tuple(fallback_documents),
        warnings=tuple(warnings),
    )


def load_semantic_documents(path: str | Path) -> list[dict[str, Any]]:
    """Load and verify a semantic artifact's manifest and documents."""

    root = Path(path).resolve()
    manifest_path = root / "manifest.json"
    if not manifest_path.is_file():
        raise FileNotFoundError(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, Mapping):
        raise ValueError("semantic manifest must be a JSON object")
    if manifest.get("schema_version") != SEMANTIC_SCHEMA_VERSION:
        raise ValueError("unsupported semantic artifact schema")
    provenance_digest = manifest.get("source_provenance_sha256")
    if not isinstance(provenance_digest, str) or not _DIGEST_RE.fullmatch(
        provenance_digest
    ):
        raise ValueError("semantic source_provenance_sha256 must be a SHA-256 digest")
    documents_name = manifest.get("documents_file")
    if (
        not isinstance(documents_name, str)
        or not documents_name
        or "/" in documents_name
        or "\\" in documents_name
        or Path(documents_name).name != documents_name
        or Path(documents_name).suffix.casefold() != ".jsonl"
    ):
        raise ValueError("semantic documents_file must be a local .jsonl filename")
    documents_path = root / documents_name
    raw = documents_path.read_bytes()
    documents_digest = manifest.get("documents_sha256")
    if not isinstance(documents_digest, str) or not _DIGEST_RE.fullmatch(
        documents_digest
    ):
        raise ValueError("semantic documents_sha256 must be a SHA-256 digest")
    actual_documents_digest = sha256(raw).hexdigest()
    if actual_documents_digest != documents_digest.casefold():
        raise ValueError(
            "semantic documents SHA-256 mismatch: "
            f"expected {documents_digest.casefold()}, got {actual_documents_digest}"
        )

    document_count = manifest.get("document_count")
    if (
        isinstance(document_count, bool)
        or not isinstance(document_count, int)
        or document_count < 0
    ):
        raise ValueError("semantic document_count must be a non-negative integer")

    documents: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    for line_number, line in enumerate(raw.decode("utf-8").splitlines(), start=1):
        if not line.strip():
            continue
        value = json.loads(line)
        if not isinstance(value, Mapping):
            raise ValueError(f"semantic document line {line_number} is not an object")
        document = dict(value)
        document_id = document.get("id")
        if not isinstance(document_id, str) or not _safe_document_id(document_id):
            raise ValueError(f"unsafe document id: {document_id!r}")
        collision_key = unicodedata.normalize("NFC", document_id).casefold()
        if collision_key in seen_ids:
            raise ValueError(f"duplicate semantic document id: {document_id!r}")
        seen_ids.add(collision_key)

        title = document.get("title")
        if (
            not isinstance(title, str)
            or not title.strip()
            or "\n" in title
            or "\r" in title
        ):
            raise ValueError(f"invalid semantic document title: {document_id}")
        body = document.get("body_markdown")
        if not isinstance(body, str):
            raise ValueError(f"semantic document body_markdown must be a string: {document_id}")
        content_digest = document.get("content_sha256")
        if not isinstance(content_digest, str) or not _DIGEST_RE.fullmatch(
            content_digest
        ):
            raise ValueError(f"semantic document content_sha256 is invalid: {document_id}")
        if sha256(body.encode("utf-8")).hexdigest() != content_digest.casefold():
            raise ValueError(
                f"semantic document content SHA-256 mismatch: {document_id}"
            )
        supplied = document.get("citations", [])
        if (
            not isinstance(supplied, list)
            or not supplied
            or not all(isinstance(item, str) and item for item in supplied)
        ):
            raise ValueError(
                f"semantic document citations must be a non-empty array of strings: {document_id}"
            )
        embedded = {match.group(0) for match in CITATION_RE.finditer(body)}
        undeclared = sorted(embedded - set(supplied))
        if undeclared:
            raise ValueError(
                f"semantic document body has undeclared citation(s) for {document_id}: "
                + ", ".join(undeclared)
            )
        documents.append(document)
    if len(documents) != document_count:
        raise ValueError("semantic document count does not match manifest")
    return documents


def _safe_document_id(value: str) -> bool:
    if (
        not value
        or len(value) > 128
        or value in {".", ".."}
        or value.endswith(".")
        or value.casefold() == "index"
        or not _DOCUMENT_ID_RE.fullmatch(value)
    ):
        return False
    return value.split(".", 1)[0].casefold() not in _WINDOWS_RESERVED_IDS


__all__ = [
    "SEMANTIC_SCHEMA_VERSION",
    "SemanticConfig",
    "SemanticDocument",
    "SemanticExport",
    "build_semantic_output",
    "load_semantic_documents",
]
