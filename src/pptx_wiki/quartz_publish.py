"""Deterministically materialize an integrated collection for Quartz.

This module is deliberately model-free.  It accepts only already-integrated,
content-addressed artifacts, validates their complete lineage, and publishes a
self-contained Quartz ``content`` tree in one directory rename.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import stat
import tempfile
import unicodedata
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from hashlib import sha256
from html import escape as html_escape
from html import unescape as html_unescape
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Any
from urllib.parse import quote, unquote, urlsplit

from .source_semantic import canonical_pr_number, load_source_semantic

QUARTZ_SCHEMA_VERSION = "pptx-wiki.quartz.v1"
INTEGRATED_SCHEMA_VERSION = "pptx-wiki.integrated.v1"
COLLECTION_SOURCE_SCHEMA_VERSION = "pptx-wiki.collection-source.v1"

_INTEGRATED_FILES = (
    "source-map.jsonl",
    "entities.jsonl",
    "pages.jsonl",
    "coverage.jsonl",
)
_DIGEST_RE = re.compile(r"^[0-9a-fA-F]{64}$")
_SAFE_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
_QUALIFIED_CITATION_RE = re.compile(
    r"\[@(?P<source>[A-Za-z0-9][A-Za-z0-9._-]{0,127})/"
    r"slide-(?P<slide>[1-9]\d*)#(?P<element>[^\]\s#]+)\]"
)
_LOCAL_CITATION_RE = re.compile(
    r"\[slide-(?P<slide>[1-9]\d*)#(?P<element>[^\]\s#]+)\]"
)
_ANCHOR_RE = re.compile(
    r"<a\s+[^>]*?id\s*=\s*([\"'])(?P<id>.*?)\1[^>]*?>",
    re.IGNORECASE,
)
_INLINE_LINK_RE = re.compile(
    r"(?P<image>!)?\[(?P<label>(?:\\.|[^\]])*)\]"
    r"\((?P<target>[^)\n]+)\)"
)
_REFERENCE_LINK_RE = re.compile(
    r"^(?P<prefix>\s{0,3}\[[^\]\r\n]+\]:\s*)"
    r"(?P<target><[^>\r\n]+>|\S+)(?P<rest>[^\r\n]*)$",
    re.MULTILINE,
)
_HTML_LINK_RE = re.compile(
    r"(?P<prefix>\b(?P<attribute>src|href)\s*=\s*)"
    r"(?P<quote>[\"'])(?P<target>.*?)(?P=quote)",
    re.IGNORECASE,
)
_WIKILINK_RE = re.compile(r"\[\[(?P<value>[^\[\]\r\n]+)\]\]")
_HEADING_RE = re.compile(r"^(?P<level>#{1,6})\s+(?P<title>.+?)\s*#*\s*$")
_NUMBER_RE = re.compile(
    r"(?<![0-9A-Za-z])[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?(?![0-9A-Za-z])"
)
_IDENTIFIER_TOKEN_RE = re.compile(
    r"(?i)(?<![A-Z0-9])(?:[A-Z]{2,8}[ \t_-]+[A-Z0-9][A-Z0-9._/\-]{0,63})(?![A-Z0-9])"
)
_WINDOWS_RESERVED = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{number}" for number in range(1, 10)),
    *(f"lpt{number}" for number in range(1, 10)),
}
_ENTITY_TYPES = {
    "person",
    "organization",
    "product",
    "system",
    "project",
    "location",
    "concept",
    "metric",
    "other",
}
_OUTPUT_DIRECTORIES = (
    "content",
    "content/topics",
    "content/entities",
    "content/prs",
    "content/sources",
    "content/evidence",
    "content/assets",
)


@dataclass(frozen=True, slots=True)
class QuartzExport:
    output_dir: Path
    content_dir: Path
    index_path: Path
    manifest_path: Path
    readme_path: Path
    page_paths: tuple[Path, ...]
    asset_paths: tuple[Path, ...]
    page_count: int
    source_count: int
    entity_count: int
    pr_count: int

    @property
    def report_path(self) -> Path:
        """Compatibility-style alias for the publication manifest."""

        return self.manifest_path


@dataclass(frozen=True, slots=True)
class _Source:
    source_id: str
    pr_numbers: tuple[str, ...]
    root: Path
    source_name: str
    source_sha256: str
    semantic_manifest: Mapping[str, Any]
    semantic_markdown: str
    slides: Mapping[int, tuple[Path, str]]
    provenance_records: Mapping[str, Mapping[str, Any]]
    asset_digests: Mapping[str, str]


@dataclass(frozen=True, slots=True)
class _Citation:
    qualified: str
    source_id: str
    pr_numbers: tuple[str, ...]
    pr_variants: tuple[str, ...]
    local: str
    slide_number: int
    element_id: str
    content_sha256: str
    numeric_tokens: tuple[str, ...]
    identifier_tokens: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class _Entity:
    entity_id: str
    entity_type: str
    canonical_name: str
    aliases: tuple[str, ...]
    description: str
    citations: tuple[str, ...]
    source_ids: tuple[str, ...]
    pr_numbers: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class _Page:
    page_id: str
    kind: str
    title: str
    description: str
    body_markdown: str
    citations: tuple[str, ...]
    source_ids: tuple[str, ...]
    pr_numbers: tuple[str, ...]
    entity_ids: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class _Coverage:
    source_id: str
    semantic_document_id: str
    pr_numbers: tuple[str, ...]
    page_ids: tuple[str, ...]
    entity_ids: tuple[str, ...]


class _OutputFiles:
    def __init__(self) -> None:
        self.values: dict[str, bytes] = {}
        self._collision_keys: dict[str, str] = {}

    def add_text(self, path: str, value: str) -> None:
        self.add_bytes(path, value.encode("utf-8"))

    def add_bytes(self, path: str, value: bytes) -> None:
        normalized = _safe_relative_posix(path, "output path")
        collision_key = unicodedata.normalize("NFC", normalized).casefold()
        previous = self._collision_keys.get(collision_key)
        if previous is not None and previous != normalized:
            raise ValueError(
                f"Quartz output path collision: {previous!r} and {normalized!r}"
            )
        if normalized in self.values:
            if self.values[normalized] != value:
                raise ValueError(f"conflicting Quartz output bytes for {normalized}")
            return
        self._collision_keys[collision_key] = normalized
        self.values[normalized] = value


class _AssetStore:
    def __init__(self, files: _OutputFiles) -> None:
        self.files = files

    def add(self, source: _Source, asset_path: Path) -> str:
        resolved = asset_path.resolve(strict=True)
        if resolved.is_symlink() or not resolved.is_file():
            raise ValueError(f"asset is not a regular file: {asset_path}")
        raw = resolved.read_bytes()
        digest = sha256(raw).hexdigest()
        expected = source.asset_digests.get(_asset_key(resolved))
        if expected is None:
            raise ValueError(
                f"local asset is not declared by parsed provenance for "
                f"{source.source_id}: {asset_path}"
            )
        if digest != expected:
            raise ValueError(
                f"source asset SHA-256 mismatch for {source.source_id}: "
                f"expected {expected}, got {digest}"
            )
        suffix = resolved.suffix.casefold()
        if not re.fullmatch(r"\.[a-z0-9]{1,10}", suffix):
            suffix = ".bin"
        output_path = f"content/assets/{source.source_id}/{digest}{suffix}"
        self.files.add_bytes(output_path, raw)
        return output_path


def publish_quartz(
    collection_dir: str | Path,
    integrated_dir: str | Path,
    output_dir: str | Path,
    *,
    site_title: str = "Reliability Analysis Wiki",
) -> QuartzExport:
    """Publish a validated integrated artifact as self-contained Quartz input.

    No model is called.  Validation and rendering finish before a sibling
    staging directory is created, and the completed tree is committed with one
    directory rename.  ``output_dir`` must be absent or an existing empty
    directory.
    """

    title = _one_line(site_title, "site_title")
    # Keep the lexical paths until every existing component has been checked.
    # Resolving first would hide a symlink or Windows junction from the guards.
    collection_input = Path(collection_dir).expanduser().absolute()
    integrated_input = Path(integrated_dir).expanduser().absolute()
    destination_input = Path(output_dir).expanduser().absolute()
    _reject_reparse_components(collection_input, "collection")
    _reject_reparse_components(integrated_input, "integrated")
    _reject_reparse_components(destination_input, "Quartz output")
    collection = collection_input.resolve(strict=True)
    integrated = integrated_input.resolve(strict=True)
    destination = destination_input.resolve(strict=False)
    _validate_input_root(collection, "collection")
    _validate_input_root(integrated, "integrated")
    _ensure_empty_destination(destination)

    manifest_path = integrated / "manifest.json"
    manifest_raw = _read_regular_file(manifest_path, integrated, "integrated manifest")
    manifest = _json_object(manifest_raw, "integrated manifest")
    artifacts = _load_integrated_artifacts(integrated, manifest)
    sources = _load_sources(collection, manifest)
    citations = _parse_source_map(artifacts["source-map.jsonl"], sources)
    entities = _parse_entities(artifacts["entities.jsonl"], citations, sources)
    pages = _parse_pages(artifacts["pages.jsonl"], citations, entities, sources)
    coverage = _parse_coverage(
        artifacts["coverage.jsonl"], sources, pages, entities
    )
    _validate_coverage(manifest, sources, pages, entities, coverage)

    files = _OutputFiles()
    assets = _AssetStore(files)
    evidence_anchors = _validate_citation_evidence(citations, sources)
    citation_links = {
        token: _citation_wikilink(citation) for token, citation in citations.items()
    }

    for source_id in sorted(sources, key=_stable_key):
        source = sources[source_id]
        for slide_number in sorted(source.slides):
            slide_path, slide_markdown = source.slides[slide_number]
            output_path = (
                f"content/evidence/{source_id}/slide-{slide_number:04d}.md"
            )
            body = _prepare_input_markdown(
                slide_markdown,
                label=f"evidence slide {source_id}/{slide_number}",
                allow_frontmatter=True,
            )
            body = _replace_citations(
                body,
                citations,
                citation_links,
                local_source_id=source_id,
                allow_unmapped_local=True,
            )
            body = _rewrite_local_links(
                body,
                input_path=slide_path,
                output_path=output_path,
                source=source,
                assets=assets,
            )
            evidence_title = f"{source_id} · Slide {slide_number}"
            rendered = _render_markdown_page(
                title=evidence_title,
                description=f"Source-faithful evidence for slide {slide_number}.",
                aliases=(),
                tags=("evidence", f"source/{source_id}"),
                pr_numbers=source.pr_numbers,
                source_ids=(source_id,),
                body=body,
            )
            files.add_text(output_path, rendered)

        source_output = f"content/sources/{source_id}.md"
        semantic_body = _prepare_input_markdown(
            source.semantic_markdown,
            label=f"semantic Markdown for {source_id}",
            allow_frontmatter=True,
        )
        semantic_body = _replace_citations(
            semantic_body,
            citations,
            citation_links,
            local_source_id=source_id,
            allow_unmapped_local=False,
        )
        semantic_body = _rewrite_local_links(
            semantic_body,
            input_path=source.root / "semantic" / "semantic.md",
            output_path=source_output,
            source=source,
            assets=assets,
        )
        source_heading = " / ".join(source.pr_numbers)
        source_body = _with_heading(
            f"{source_heading} · {source.source_name}", semantic_body
        )
        source_description = f"Validated semantic source for {source_heading}."
        files.add_text(
            source_output,
            _render_markdown_page(
                title=f"{source_heading} · {source.source_name}",
                description=source_description,
                aliases=(),
                tags=("source",),
                pr_numbers=source.pr_numbers,
                source_ids=(source_id,),
                body=source_body,
            ),
        )

    for entity_id in sorted(entities, key=_stable_key):
        entity = entities[entity_id]
        body = _render_entity_body(entity, citation_links)
        files.add_text(
            f"content/entities/{entity_id}.md",
            _render_markdown_page(
                title=entity.canonical_name,
                description=_frontmatter_description(entity.description),
                aliases=entity.aliases,
                tags=("entity", f"entity/{_tag_token(entity.entity_type)}"),
                pr_numbers=entity.pr_numbers,
                source_ids=entity.source_ids,
                body=body,
            ),
        )

    for page_id in sorted(pages, key=_stable_key):
        page = pages[page_id]
        body = _replace_citations(
            _prepare_input_markdown(
                page.body_markdown,
                label=f"integrated page {page_id}",
                allow_frontmatter=False,
            ),
            citations,
            citation_links,
        )
        description_body = _replace_citations(
            page.description, citations, citation_links
        ).strip()
        if description_body and description_body not in body:
            body = description_body + "\n\n" + body
        body = _append_entity_links(body, page.entity_ids, entities)
        body = _with_heading(page.title, body)
        files.add_text(
            f"content/topics/{page_id}.md",
            _render_markdown_page(
                title=page.title,
                description=_frontmatter_description(page.description),
                aliases=(),
                tags=("topic", f"topic/{_tag_token(page.kind)}"),
                pr_numbers=page.pr_numbers,
                source_ids=page.source_ids,
                body=body,
            ),
        )

    pr_slugs = _assign_pr_slugs(
        {pr for source in sources.values() for pr in source.pr_numbers}
    )
    for pr_number in sorted(pr_slugs, key=_stable_key):
        slug = pr_slugs[pr_number]
        body = _render_pr_body(pr_number, sources, pages, entities)
        pr_sources = tuple(
            source_id
            for source_id in sorted(sources, key=_stable_key)
            if pr_number in sources[source_id].pr_numbers
        )
        files.add_text(
            f"content/prs/{slug}.md",
            _render_markdown_page(
                title=pr_number,
                description=f"Canonical source index for {pr_number}.",
                aliases=(),
                tags=("pr",),
                pr_numbers=(pr_number,),
                source_ids=pr_sources,
                body=body,
            ),
        )

    index_body = _render_index_body(title, sources, pages, entities, pr_slugs)
    all_prs = tuple(sorted(pr_slugs, key=_stable_key))
    all_sources = tuple(sorted(sources, key=_stable_key))
    files.add_text(
        "content/index.md",
        _render_markdown_page(
            title=title,
            description="Integrated, evidence-linked reliability analysis knowledge base.",
            aliases=(),
            tags=("index",),
            pr_numbers=all_prs,
            source_ids=all_sources,
            body=index_body,
        ),
    )
    files.add_text("README.md", _render_readme(title))

    _validate_rendered_tree(files.values, evidence_anchors)
    quartz_manifest = _render_quartz_manifest(
        site_title=title,
        integrated_manifest_sha256=sha256(manifest_raw).hexdigest(),
        sources=sources,
        entities=entities,
        pages=pages,
        pr_slugs=pr_slugs,
        files=files.values,
    )
    files.add_text("quartz-manifest.json", quartz_manifest)

    _commit_files(destination, files.values)
    page_relpaths = tuple(
        path
        for path in sorted(files.values, key=_stable_key)
        if path.startswith("content/") and path.endswith(".md")
    )
    asset_relpaths = tuple(
        path
        for path in sorted(files.values, key=_stable_key)
        if path.startswith("content/assets/")
    )
    return QuartzExport(
        output_dir=destination,
        content_dir=destination / "content",
        index_path=destination / "content" / "index.md",
        manifest_path=destination / "quartz-manifest.json",
        readme_path=destination / "README.md",
        page_paths=tuple(destination / PurePosixPath(path) for path in page_relpaths),
        asset_paths=tuple(destination / PurePosixPath(path) for path in asset_relpaths),
        page_count=len(page_relpaths),
        source_count=len(sources),
        entity_count=len(entities),
        pr_count=len(pr_slugs),
    )


def _validate_input_root(path: Path, label: str) -> None:
    if not path.is_dir():
        raise FileNotFoundError(f"{label} directory not found: {path}")
    if path.is_symlink():
        raise ValueError(f"{label} directory must not be a symlink: {path}")


def _reject_reparse_components(path: Path, label: str) -> None:
    """Reject symlinks and Windows reparse points in a lexical path."""

    current = path
    existing: list[Path] = []
    while True:
        if current.exists() or current.is_symlink():
            existing.append(current)
        if current.parent == current:
            break
        current = current.parent
    for component in reversed(existing):
        if _is_reparse_point(component):
            raise ValueError(
                f"{label} path must not contain a symlink/junction/reparse point: "
                f"{component}"
            )


def _is_reparse_point(path: Path) -> bool:
    if path.is_symlink():
        return True
    is_junction = getattr(path, "is_junction", None)
    if callable(is_junction) and is_junction():
        return True
    try:
        attributes = getattr(os.lstat(path), "st_file_attributes", 0)
    except OSError:
        return False
    return bool(attributes & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))


def _ensure_empty_destination(path: Path) -> None:
    if path.exists() and not path.is_dir():
        raise ValueError(f"Quartz output exists and is not a directory: {path}")
    if path.is_symlink():
        raise ValueError(f"Quartz output must not be a symlink: {path}")
    if path.exists() and any(path.iterdir()):
        raise ValueError(f"Quartz output directory is not empty: {path}")
    parent = path.parent
    if not parent.is_dir():
        raise FileNotFoundError(f"Quartz output parent does not exist: {parent}")
    if parent.is_symlink():
        raise ValueError(f"Quartz output parent must not be a symlink: {parent}")


def _read_regular_file(path: Path, root: Path, label: str) -> bytes:
    try:
        resolved = path.resolve(strict=True)
    except FileNotFoundError:
        raise FileNotFoundError(f"{label} not found: {path}") from None
    root_resolved = root.resolve(strict=True)
    if not resolved.is_relative_to(root_resolved):
        raise ValueError(f"{label} escapes its artifact root: {path}")
    if path.is_symlink() or not resolved.is_file():
        raise ValueError(f"{label} is not a regular file: {path}")
    return resolved.read_bytes()


def _json_object(raw: bytes, label: str) -> Mapping[str, Any]:
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid JSON in {label}: {exc}") from exc
    if not isinstance(value, Mapping):
        raise TypeError(f"{label} must be a JSON object")
    return value


def _jsonl_objects(raw: bytes, label: str) -> list[Mapping[str, Any]]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError(f"{label} is not valid UTF-8: {exc}") from exc
    records: list[Mapping[str, Any]] = []
    for line_number, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"invalid JSON in {label} line {line_number}: {exc}"
            ) from exc
        if not isinstance(value, Mapping):
            raise TypeError(f"{label} line {line_number} must be a JSON object")
        records.append(value)
    return records


def _load_integrated_artifacts(
    root: Path, manifest: Mapping[str, Any]
) -> dict[str, list[Mapping[str, Any]]]:
    if manifest.get("schema_version") != INTEGRATED_SCHEMA_VERSION:
        raise ValueError(
            "unsupported integrated schema: "
            f"{manifest.get('schema_version')!r}; expected {INTEGRATED_SCHEMA_VERSION!r}"
        )
    if manifest.get("coverage_complete") is not True:
        raise ValueError("integrated manifest coverage_complete must be true")
    if not isinstance(manifest.get("backend"), Mapping):
        raise TypeError("integrated manifest backend must be an object")
    if not isinstance(manifest.get("config"), Mapping):
        raise TypeError("integrated manifest config must be an object")
    warnings = manifest.get("warnings")
    if not isinstance(warnings, list) or not all(
        isinstance(item, str) for item in warnings
    ):
        raise ValueError("integrated manifest warnings must be an array of strings")

    file_entries = manifest.get("files")
    if not isinstance(file_entries, Mapping):
        raise TypeError("integrated manifest files must be an object")
    missing = [name for name in _INTEGRATED_FILES if name not in file_entries]
    if missing:
        raise ValueError(
            "integrated manifest is missing file declaration(s): "
            + ", ".join(missing)
        )

    loaded_raw: dict[str, bytes] = {}
    loaded_records: dict[str, list[Mapping[str, Any]]] = {}
    for declared_name, entry in file_entries.items():
        if not isinstance(declared_name, str):
            raise TypeError("integrated manifest file names must be strings")
        safe_name = _safe_relative_posix(declared_name, "integrated file name")
        if not isinstance(entry, Mapping):
            raise TypeError(
                f"integrated manifest files[{declared_name!r}] must be an object"
            )
        digest = _digest(entry.get("sha256"), f"files[{declared_name!r}].sha256")
        count = _non_negative_int(
            entry.get("count"), f"files[{declared_name!r}].count"
        )
        raw = _read_regular_file(root / PurePosixPath(safe_name), root, declared_name)
        actual = sha256(raw).hexdigest()
        if actual != digest:
            raise ValueError(
                f"integrated file SHA-256 mismatch for {declared_name}: "
                f"expected {digest}, got {actual}"
            )
        records = _jsonl_objects(raw, f"integrated {declared_name}")
        if len(records) != count:
            raise ValueError(
                f"integrated file count mismatch for {declared_name}: "
                f"expected {count}, got {len(records)}"
            )
        loaded_raw[declared_name] = raw
        loaded_records[declared_name] = records

    # The four contract files are fixed names.  Extra content-addressed JSONL
    # audit files are accepted only after receiving the same hash/count checks.
    return {name: loaded_records[name] for name in _INTEGRATED_FILES}


def _load_sources(
    collection: Path, manifest: Mapping[str, Any]
) -> dict[str, _Source]:
    values = manifest.get("sources")
    if not isinstance(values, list) or not values:
        raise ValueError("integrated manifest sources must be a non-empty array")
    sources_root = collection / "sources"
    if not sources_root.is_dir() or sources_root.is_symlink():
        raise FileNotFoundError(f"collection sources directory not found: {sources_root}")

    sources: dict[str, _Source] = {}
    collision_keys: set[str] = set()
    for index, value in enumerate(values):
        label = f"integrated manifest source {index}"
        if not isinstance(value, Mapping):
            raise TypeError(f"{label} must be an object")
        source_id = _safe_id(value.get("source_id"), f"{label}.source_id")
        collision_key = unicodedata.normalize("NFC", source_id).casefold()
        if collision_key in collision_keys:
            raise ValueError(f"duplicate integrated source id: {source_id}")
        collision_keys.add(collision_key)
        pr_numbers = _string_tuple(
            value.get("pr_numbers"), f"{label}.pr_numbers", nonempty=True, pr=True
        )
        canonical_keys = tuple(canonical_pr_number(item) for item in pr_numbers)
        if len(set(canonical_keys)) != len(canonical_keys):
            raise ValueError(
                f"{label}.pr_numbers repeats a canonical PR display value"
            )
        semantic_manifest_digest = _digest(
            value.get("semantic_manifest_sha256"),
            f"{label}.semantic_manifest_sha256",
        )
        semantic_markdown_digest = _digest(
            value.get("semantic_markdown_sha256"),
            f"{label}.semantic_markdown_sha256",
        )

        source_root = sources_root / source_id
        if not source_root.is_dir() or source_root.is_symlink():
            raise FileNotFoundError(f"collection source directory not found: {source_root}")
        source_raw = _read_regular_file(
            source_root / "source.json", source_root, f"source metadata for {source_id}"
        )
        source_value = _json_object(source_raw, f"source metadata for {source_id}")
        source_name, source_sha = _validate_source_metadata(
            source_value, source_id, pr_numbers
        )

        semantic_root = source_root / "semantic"
        semantic_manifest_raw = _read_regular_file(
            semantic_root / "manifest.json",
            source_root,
            f"semantic manifest for {source_id}",
        )
        actual_manifest_digest = sha256(semantic_manifest_raw).hexdigest()
        if actual_manifest_digest != semantic_manifest_digest:
            raise ValueError(
                f"semantic manifest SHA-256 mismatch for {source_id}: "
                f"expected {semantic_manifest_digest}, got {actual_manifest_digest}"
            )
        semantic_manifest = _json_object(
            semantic_manifest_raw, f"semantic manifest for {source_id}"
        )
        loaded_semantic = load_source_semantic(semantic_root)
        if loaded_semantic.get("manifest") != dict(semantic_manifest):
            raise ValueError(f"source semantic loader disagreement for {source_id}")
        semantic_raw = _read_regular_file(
            semantic_root / "semantic.md",
            source_root,
            f"semantic Markdown for {source_id}",
        )
        actual_markdown_digest = sha256(semantic_raw).hexdigest()
        if actual_markdown_digest != semantic_markdown_digest:
            raise ValueError(
                f"semantic Markdown SHA-256 mismatch for {source_id}: "
                f"expected {semantic_markdown_digest}, got {actual_markdown_digest}"
            )
        try:
            semantic_markdown = semantic_raw.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise ValueError(
                f"semantic Markdown for {source_id} is not valid UTF-8: {exc}"
            ) from exc

        parsed_root = source_root / "parsed"
        corpus_root = parsed_root / "corpus"
        slides_root = corpus_root / "slides"
        if not slides_root.is_dir() or slides_root.is_symlink():
            raise FileNotFoundError(f"parsed slides directory not found: {slides_root}")
        slides: dict[int, tuple[Path, str]] = {}
        for slide_path in sorted(slides_root.glob("slide-*.md"), key=lambda item: item.name):
            match = re.fullmatch(r"slide-(?P<number>\d+)\.md", slide_path.name)
            if match is None:
                raise ValueError(f"unsafe parsed slide filename: {slide_path.name}")
            slide_number = int(match.group("number"))
            if slide_number <= 0:
                raise ValueError(f"parsed slide number must be positive: {slide_path.name}")
            raw = _read_regular_file(
                slide_path, source_root, f"parsed slide {source_id}/{slide_number}"
            )
            try:
                markdown = raw.decode("utf-8")
            except UnicodeDecodeError as exc:
                raise ValueError(
                    f"parsed slide {source_id}/{slide_number} is not UTF-8: {exc}"
                ) from exc
            if slide_number in slides:
                raise ValueError(f"duplicate parsed slide number for {source_id}: {slide_number}")
            slides[slide_number] = (slide_path, markdown)
        if not slides:
            raise ValueError(f"collection source {source_id} has no parsed slides")

        provenance_path = corpus_root / "provenance.jsonl"
        provenance_records = _load_provenance_records(
            provenance_path, source_root, source_id
        )
        asset_digests = _load_source_asset_digests(
            provenance_records, source_root, source_id
        )
        _validate_semantic_manifest(
            semantic_manifest,
            semantic_root,
            corpus_root,
            source_root,
            source_id,
            pr_numbers,
            semantic_raw,
        )
        sources[source_id] = _Source(
            source_id=source_id,
            pr_numbers=pr_numbers,
            root=source_root,
            source_name=source_name,
            source_sha256=source_sha,
            semantic_manifest=semantic_manifest,
            semantic_markdown=semantic_markdown,
            slides=slides,
            provenance_records=provenance_records,
            asset_digests=asset_digests,
        )

    actual_directories = {
        path.name
        for path in sources_root.iterdir()
        if path.is_dir() and not path.is_symlink()
    }
    if actual_directories != set(sources):
        missing = sorted(actual_directories - set(sources), key=_stable_key)
        absent = sorted(set(sources) - actual_directories, key=_stable_key)
        details: list[str] = []
        if missing:
            details.append("unmanifested source directories: " + ", ".join(missing))
        if absent:
            details.append("missing source directories: " + ", ".join(absent))
        raise ValueError("collection source set mismatch: " + "; ".join(details))
    return sources


def _validate_source_metadata(
    value: Mapping[str, Any], source_id: str, pr_numbers: tuple[str, ...]
) -> tuple[str, str]:
    if value.get("schema_version") != COLLECTION_SOURCE_SCHEMA_VERSION:
        raise ValueError(f"unsupported collection source schema for {source_id}")
    if value.get("source_id") != source_id:
        raise ValueError(f"source.json source_id mismatch for {source_id}")
    supplied_prs = _string_tuple(
        value.get("pr_numbers"),
        f"source.json pr_numbers for {source_id}",
        nonempty=True,
        pr=True,
    )
    if supplied_prs != pr_numbers:
        raise ValueError(f"source.json PR inventory mismatch for {source_id}")
    source_sha = _digest(value.get("source_sha256"), f"source SHA-256 for {source_id}")
    source_name = _one_line(value.get("source_name"), f"source name for {source_id}")
    if Path(source_name).name != source_name or PureWindowsPath(source_name).name != source_name:
        raise ValueError(f"source_name must be a filename for {source_id}")
    occurrences = value.get("occurrences")
    if not isinstance(occurrences, list) or not occurrences:
        raise ValueError(f"source occurrences must be a non-empty array for {source_id}")
    for index, occurrence in enumerate(occurrences):
        label = f"source occurrence {source_id}/{index}"
        if not isinstance(occurrence, Mapping):
            raise TypeError(f"{label} must be an object")
        # New collection artifacts expose only a non-reversible occurrence ID.
        # A legacy private ``path`` is accepted for input validation but never
        # copied into public output.
        if "occurrence_id" in occurrence:
            occurrence_id = _one_line(
                occurrence.get("occurrence_id"), f"{label}.occurrence_id"
            )
            if re.fullmatch(r"occurrence-[0-9a-f]{16,64}", occurrence_id) is None:
                raise ValueError(f"{label}.occurrence_id is invalid")
        elif "path" in occurrence:
            _one_line(occurrence.get("path"), f"{label}.path")
        else:
            raise ValueError(f"{label} requires occurrence_id")
        name = _one_line(occurrence.get("name"), f"{label}.name")
        if Path(name).name != name or PureWindowsPath(name).name != name:
            raise ValueError(f"{label}.name must be a filename")
        _non_negative_int(occurrence.get("size"), f"{label}.size")
        _non_negative_int(occurrence.get("mtime_ns"), f"{label}.mtime_ns")
        occurrence_sha = _digest(occurrence.get("sha256"), f"{label}.sha256")
        if occurrence_sha != source_sha:
            raise ValueError(f"{label}.sha256 does not match source_sha256")
    return source_name, source_sha


def _load_provenance_records(
    path: Path, source_root: Path, source_id: str
) -> dict[str, Mapping[str, Any]]:
    if not path.exists():
        return {}
    raw = _read_regular_file(path, source_root, f"parsed provenance for {source_id}")
    records = _jsonl_objects(raw, f"parsed provenance for {source_id}")
    result: dict[str, Mapping[str, Any]] = {}
    for index, record in enumerate(records):
        label = f"parsed provenance {source_id}/{index}"
        citation = _one_line(record.get("citation"), f"{label}.citation")
        match = _LOCAL_CITATION_RE.fullmatch(citation)
        if match is None:
            raise ValueError(f"invalid local citation in {label}: {citation!r}")
        content = record.get("content")
        if not isinstance(content, str):
            raise TypeError(f"{label}.content must be a string")
        actual = sha256(content.encode("utf-8")).hexdigest()
        supplied = _digest(record.get("content_sha256"), f"{label}.content_sha256")
        if actual != supplied:
            raise ValueError(f"content SHA-256 mismatch in {label}")
        slide_number = _positive_int(record.get("slide_number"), f"{label}.slide_number")
        element_id = _element_id(record.get("element_id"), f"{label}.element_id")
        if slide_number != int(match.group("slide")) or element_id != match.group("element"):
            raise ValueError(f"local citation fields disagree in {label}")
        if citation in result:
            raise ValueError(f"duplicate parsed provenance citation for {source_id}: {citation}")
        result[citation] = dict(record)
    return result


def _load_source_asset_digests(
    records: Mapping[str, Mapping[str, Any]],
    source_root: Path,
    source_id: str,
) -> dict[str, str]:
    """Authenticate every provenance-declared source asset before rendering."""

    assets_root = source_root / "parsed" / "source-assets"
    result: dict[str, str] = {}
    for citation, record in records.items():
        raw_path = record.get("asset_path")
        if raw_path is None:
            continue
        if not isinstance(raw_path, str) or not raw_path.strip():
            raise TypeError(
                f"parsed provenance asset_path must be a non-empty string: "
                f"{source_id}/{citation}"
            )
        if any(character in raw_path for character in ("\x00", "\r", "\n")):
            raise ValueError(
                f"parsed provenance asset_path contains control characters: "
                f"{source_id}/{citation}"
            )
        resolved = _resolve_provenance_asset(
            raw_path,
            source_root=source_root,
            assets_root=assets_root,
            label=f"{source_id}/{citation}",
        )
        metadata = record.get("metadata")
        if not isinstance(metadata, Mapping):
            raise TypeError(
                f"parsed provenance asset metadata must be an object: "
                f"{source_id}/{citation}"
            )
        expected = _digest(
            metadata.get("image_sha256"),
            f"parsed provenance image_sha256 for {source_id}/{citation}",
        )
        raw = _read_regular_file(
            resolved,
            source_root,
            f"parsed source asset for {source_id}/{citation}",
        )
        actual = sha256(raw).hexdigest()
        if actual != expected:
            raise ValueError(
                f"source asset SHA-256 mismatch for {source_id}/{citation}: "
                f"expected {expected}, got {actual}"
            )
        image_bytes = metadata.get("image_bytes")
        if image_bytes is not None and _non_negative_int(
            image_bytes,
            f"parsed provenance image_bytes for {source_id}/{citation}",
        ) != len(raw):
            raise ValueError(
                f"source asset byte count mismatch for {source_id}/{citation}"
            )
        key = _asset_key(resolved)
        previous = result.get(key)
        if previous is not None and previous != expected:
            raise ValueError(
                f"conflicting source asset digests for {source_id}: {raw_path}"
            )
        result[key] = expected
    return result


def _resolve_provenance_asset(
    value: str,
    *,
    source_root: Path,
    assets_root: Path,
    label: str,
) -> Path:
    windows_absolute = PureWindowsPath(value).is_absolute() or bool(
        re.match(r"^[A-Za-z]:[\\/]", value)
    )
    raw = Path(value)
    candidates = (
        (raw,)
        if windows_absolute or raw.is_absolute()
        else (
            source_root / raw,
            source_root / "parsed" / raw,
            assets_root / raw,
        )
    )
    resolved_assets_root = assets_root.resolve(strict=True)
    for candidate in candidates:
        try:
            resolved = candidate.resolve(strict=True)
        except (FileNotFoundError, OSError):
            continue
        if not resolved.is_relative_to(resolved_assets_root):
            continue
        if candidate.is_symlink() or resolved.is_symlink() or not resolved.is_file():
            raise ValueError(f"parsed provenance asset is not a regular file: {label}")
        return resolved
    raise ValueError(
        f"parsed provenance asset is missing or outside parsed/source-assets: {label}"
    )


def _asset_key(path: Path) -> str:
    return os.path.normcase(str(path.resolve(strict=True)))


def _validate_semantic_manifest(
    manifest: Mapping[str, Any],
    semantic_root: Path,
    corpus_root: Path,
    source_root: Path,
    source_id: str,
    pr_numbers: tuple[str, ...],
    semantic_raw: bytes,
) -> None:
    schema = manifest.get("schema_version")
    if schema != "pptx-wiki.source-semantic.v1":
        raise ValueError(f"unsupported semantic manifest schema for {source_id}: {schema!r}")
    identity = manifest.get("source_identity")
    if not isinstance(identity, Mapping):
        raise TypeError(f"source semantic identity is missing for {source_id}")
    if identity.get("source_id") != source_id:
        raise ValueError(f"semantic manifest source_id mismatch for {source_id}")
    identity_prs = _string_tuple(
        identity.get("pr_numbers"),
        f"semantic manifest pr_numbers for {source_id}",
        nonempty=True,
        pr=True,
    )
    if identity_prs != pr_numbers:
        raise ValueError(f"semantic manifest PR inventory mismatch for {source_id}")
    identity_digest = _digest(
        identity.get("source_sha256"),
        f"semantic manifest source digest for {source_id}",
    )
    source_metadata = _json_object(
        _read_regular_file(
            source_root / "source.json", source_root, f"source metadata for {source_id}"
        ),
        f"source metadata for {source_id}",
    )
    if identity_digest != str(source_metadata.get("source_sha256", "")).casefold():
        raise ValueError(f"semantic/source identity digest mismatch for {source_id}")
    if identity.get("source_name") != source_metadata.get("source_name"):
        raise ValueError(f"semantic/source identity name mismatch for {source_id}")

    markdown_name = _safe_relative_posix(
        manifest.get("markdown_file"), f"semantic markdown_file for {source_id}"
    )
    if markdown_name != "semantic.md":
        raise ValueError(f"source semantic markdown_file must be semantic.md for {source_id}")
    markdown_digest = _digest(
        manifest.get("markdown_sha256"), f"semantic markdown_sha256 for {source_id}"
    )
    if sha256(semantic_raw).hexdigest() != markdown_digest:
        raise ValueError(f"semantic Markdown manifest digest mismatch for {source_id}")

    provenance_digest = manifest.get("source_provenance_sha256")
    if provenance_digest is not None:
        expected = _digest(
            provenance_digest, f"source provenance digest for {source_id}"
        )
        raw = _read_regular_file(
            corpus_root / "provenance.jsonl",
            source_root,
            f"parsed provenance for {source_id}",
        )
        if sha256(raw).hexdigest() != expected:
            raise ValueError(f"source provenance digest mismatch for {source_id}")

    _validate_manifest_jsonl_pair(
        manifest,
        semantic_root,
        source_root,
        source_id,
        prefix="documents",
        count_key="document_count",
    )
    _validate_manifest_jsonl_pair(
        manifest,
        semantic_root,
        source_root,
        source_id,
        prefix="decisions",
        count_key="decision_count",
        optional=True,
    )
    ledger = manifest.get("pr_ledger")
    if not isinstance(ledger, list) or not ledger:
        raise ValueError(f"source semantic PR ledger is missing for {source_id}")
    ledger_prs: set[str] = set()
    ledger_canonical_keys: set[str] = set()
    for index, item in enumerate(ledger):
        label = f"source semantic PR ledger {source_id}/{index}"
        if not isinstance(item, Mapping):
            raise TypeError(f"{label} must be an object")
        ledger_value = _pr_number(item.get("value"), f"{label}.value")
        ledger_prs.add(ledger_value)
        supplied_key = _one_line(
            item.get("canonical_key"), f"{label}.canonical_key"
        )
        expected_key = canonical_pr_number(ledger_value)
        if supplied_key != expected_key:
            raise ValueError(f"{label}.canonical_key does not match its exact value")
        ledger_canonical_keys.add(expected_key)
        citation = _one_line(item.get("citation"), f"{label}.citation")
        if _LOCAL_CITATION_RE.fullmatch(citation) is None:
            raise ValueError(f"{label}.citation is invalid")
        _positive_int(item.get("slide_number"), f"{label}.slide_number")
        _element_id(item.get("element_id"), f"{label}.element_id")
        _digest(item.get("content_sha256"), f"{label}.content_sha256")
    source_keys = tuple(canonical_pr_number(value) for value in pr_numbers)
    if len(set(source_keys)) != len(source_keys):
        raise ValueError(
            f"source semantic PR inventory repeats a canonical PR for {source_id}"
        )
    if ledger_canonical_keys != set(source_keys) or not set(pr_numbers) <= ledger_prs:
        raise ValueError(f"source semantic PR ledger inventory mismatch for {source_id}")


def _validate_manifest_jsonl_pair(
    manifest: Mapping[str, Any],
    artifact_root: Path,
    source_root: Path,
    source_id: str,
    *,
    prefix: str,
    count_key: str,
    optional: bool = False,
) -> None:
    file_key = f"{prefix}_file"
    digest_key = f"{prefix}_sha256"
    has_any = file_key in manifest or digest_key in manifest or count_key in manifest
    if not has_any and optional:
        return
    if file_key not in manifest or digest_key not in manifest or count_key not in manifest:
        raise ValueError(
            f"semantic manifest must declare {file_key}, {digest_key}, and {count_key} "
            f"together for {source_id}"
        )
    name = _safe_relative_posix(manifest.get(file_key), f"{file_key} for {source_id}")
    if not name.casefold().endswith(".jsonl"):
        raise ValueError(f"{file_key} must name a JSONL file for {source_id}")
    digest = _digest(manifest.get(digest_key), f"{digest_key} for {source_id}")
    count = _non_negative_int(manifest.get(count_key), f"{count_key} for {source_id}")
    raw = _read_regular_file(
        artifact_root / PurePosixPath(name), source_root, f"semantic {prefix} for {source_id}"
    )
    if sha256(raw).hexdigest() != digest:
        raise ValueError(f"semantic {prefix} digest mismatch for {source_id}")
    records = _jsonl_objects(raw, f"semantic {prefix} for {source_id}")
    if len(records) != count:
        raise ValueError(f"semantic {prefix} count mismatch for {source_id}")


def _parse_source_map(
    records: Sequence[Mapping[str, Any]], sources: Mapping[str, _Source]
) -> dict[str, _Citation]:
    if not records:
        raise ValueError("integrated source-map.jsonl must not be empty")
    result: dict[str, _Citation] = {}
    local_keys: set[tuple[str, str]] = set()
    for index, value in enumerate(records):
        label = f"source-map.jsonl line {index + 1}"
        qualified = _one_line(
            value.get("qualified_citation"), f"{label}.qualified_citation"
        )
        source_id = _safe_id(value.get("source_id"), f"{label}.source_id")
        if source_id not in sources:
            raise ValueError(f"{label} references unknown source_id: {source_id}")
        pr_numbers = _string_tuple(
            value.get("pr_numbers"), f"{label}.pr_numbers", nonempty=False, pr=True
        )
        pr_variants = _string_tuple(
            value.get("pr_variants"),
            f"{label}.pr_variants",
            nonempty=False,
            pr=True,
        )
        local = _one_line(value.get("local_citation"), f"{label}.local_citation")
        slide_number = _positive_int(value.get("slide_number"), f"{label}.slide_number")
        element_id = _element_id(value.get("element_id"), f"{label}.element_id")
        content_digest = _digest(
            value.get("content_sha256"), f"{label}.content_sha256"
        )
        numeric_tokens = _string_tuple(
            value.get("numeric_tokens"),
            f"{label}.numeric_tokens",
            nonempty=False,
        )
        identifier_tokens = _identifier_token_tuple(
            value.get("identifier_tokens"),
            f"{label}.identifier_tokens",
        )
        local_match = _LOCAL_CITATION_RE.fullmatch(local)
        if (
            local_match is None
            or int(local_match.group("slide")) != slide_number
            or local_match.group("element") != element_id
        ):
            raise ValueError(f"{label} local citation fields disagree")
        qualified_match = _QUALIFIED_CITATION_RE.fullmatch(qualified)
        if (
            qualified_match is None
            or qualified_match.group("source") != source_id
            or int(qualified_match.group("slide")) != slide_number
            or qualified_match.group("element") != element_id
        ):
            raise ValueError(f"{label} qualified citation fields disagree")
        if qualified in result:
            raise ValueError(f"duplicate qualified citation: {qualified}")
        local_key = (source_id, local)
        if local_key in local_keys:
            raise ValueError(f"duplicate source/local citation mapping: {source_id} {local}")
        local_keys.add(local_key)
        provenance = sources[source_id].provenance_records.get(local)
        if provenance is None:
            raise ValueError(f"{label} has no parsed provenance record")
        visible = f"{provenance.get('slide_title', '')}\n{provenance.get('content', '')}"
        expected_numeric = tuple(sorted(_numeric_tokens(visible), key=_stable_key))
        expected_identifiers = _identifier_tokens(visible)
        if numeric_tokens != expected_numeric:
            raise ValueError(f"{label} numeric token inventory mismatch")
        if identifier_tokens != expected_identifiers:
            raise ValueError(f"{label} identifier token inventory mismatch")
        expected_variants = tuple(
            dict.fromkeys(
                str(item.get("value"))
                for item in sources[source_id].semantic_manifest.get("pr_ledger", ())
                if isinstance(item, Mapping) and item.get("citation") == local
            )
        )
        if pr_variants != expected_variants:
            raise ValueError(f"{label} citation-level PR variant inventory mismatch")
        # Integration emits citation-level PRs in the citation's local
        # first-seen order, while using the source inventory only to select
        # each canonical display spelling.  A citation may therefore contain
        # the same PR set in a different order from the source-wide ledger.
        canonical_display = {
            canonical_pr_number(value): value
            for value in sources[source_id].pr_numbers
        }
        expected_prs = tuple(
            dict.fromkeys(
                canonical_display[canonical_pr_number(value)]
                for value in expected_variants
            )
        )
        if pr_numbers != expected_prs:
            raise ValueError(f"{label} citation-level PR inventory mismatch")
        result[qualified] = _Citation(
            qualified=qualified,
            source_id=source_id,
            pr_numbers=pr_numbers,
            pr_variants=pr_variants,
            local=local,
            slide_number=slide_number,
            element_id=element_id,
            content_sha256=content_digest,
            numeric_tokens=numeric_tokens,
            identifier_tokens=identifier_tokens,
        )
    return result


def _parse_entities(
    records: Sequence[Mapping[str, Any]],
    citations: Mapping[str, _Citation],
    sources: Mapping[str, _Source],
) -> dict[str, _Entity]:
    result: dict[str, _Entity] = {}
    collision_keys: set[str] = set()
    alias_registry: dict[str, str] = {}
    for index, value in enumerate(records):
        label = f"entities.jsonl line {index + 1}"
        entity_id = _safe_id(value.get("id"), f"{label}.id")
        collision_key = unicodedata.normalize("NFC", entity_id).casefold()
        if collision_key in collision_keys:
            raise ValueError(f"duplicate entity id: {entity_id}")
        collision_keys.add(collision_key)
        entity_type = _one_line(value.get("type"), f"{label}.type")
        if entity_type not in _ENTITY_TYPES:
            raise ValueError(
                f"{label}.type must be one of: {', '.join(sorted(_ENTITY_TYPES))}"
            )
        canonical_name = _one_line(
            value.get("canonical_name"), f"{label}.canonical_name"
        )
        aliases = _string_tuple(
            value.get("aliases"), f"{label}.aliases", nonempty=False
        )
        description = _text(value.get("description"), f"{label}.description")
        declared_citations = _string_tuple(
            value.get("citations"), f"{label}.citations", nonempty=True
        )
        _validate_declared_citations(
            declared_citations, citations, f"entity {entity_id}"
        )
        _validate_body_citations(description, declared_citations, f"entity {entity_id}")
        source_ids = _string_tuple(
            value.get("source_ids"), f"{label}.source_ids", nonempty=True, ids=True
        )
        pr_numbers = _string_tuple(
            value.get("pr_numbers"), f"{label}.pr_numbers", nonempty=True, pr=True
        )
        _validate_lineage_lists(
            declared_citations,
            source_ids,
            pr_numbers,
            citations,
            f"entity {entity_id}",
            sources,
        )
        for alias in (canonical_name, *aliases):
            alias_key = unicodedata.normalize("NFC", alias).casefold()
            owner = alias_registry.get(alias_key)
            if owner is not None and owner != entity_id:
                raise ValueError(
                    f"entity alias/title collision between {owner!r} and {entity_id!r}: {alias!r}"
                )
            alias_registry[alias_key] = entity_id
        result[entity_id] = _Entity(
            entity_id=entity_id,
            entity_type=entity_type,
            canonical_name=canonical_name,
            aliases=aliases,
            description=description,
            citations=declared_citations,
            source_ids=source_ids,
            pr_numbers=pr_numbers,
        )
    return result


def _parse_pages(
    records: Sequence[Mapping[str, Any]],
    citations: Mapping[str, _Citation],
    entities: Mapping[str, _Entity],
    sources: Mapping[str, _Source],
) -> dict[str, _Page]:
    if not records:
        raise ValueError("integrated pages.jsonl must not be empty")
    result: dict[str, _Page] = {}
    collision_keys: set[str] = set()
    for index, value in enumerate(records):
        label = f"pages.jsonl line {index + 1}"
        page_id = _safe_id(value.get("id"), f"{label}.id")
        collision_key = unicodedata.normalize("NFC", page_id).casefold()
        if collision_key in collision_keys:
            raise ValueError(f"duplicate integrated page id: {page_id}")
        collision_keys.add(collision_key)
        kind = _one_line(value.get("kind"), f"{label}.kind")
        if kind != "topic":
            raise ValueError(f"{label}.kind must be 'topic'")
        title = _one_line(value.get("title"), f"{label}.title")
        description = _text(value.get("description"), f"{label}.description")
        body = _text(value.get("body_markdown"), f"{label}.body_markdown")
        declared_citations = _string_tuple(
            value.get("citations"), f"{label}.citations", nonempty=True
        )
        _validate_declared_citations(
            declared_citations, citations, f"integrated page {page_id}"
        )
        _validate_body_citations(
            description + "\n" + body,
            declared_citations,
            f"integrated page {page_id}",
        )
        source_ids = _string_tuple(
            value.get("source_ids"), f"{label}.source_ids", nonempty=True, ids=True
        )
        pr_numbers = _string_tuple(
            value.get("pr_numbers"), f"{label}.pr_numbers", nonempty=True, pr=True
        )
        _validate_lineage_lists(
            declared_citations,
            source_ids,
            pr_numbers,
            citations,
            f"integrated page {page_id}",
            sources,
        )
        entity_ids = _string_tuple(
            value.get("entity_ids"), f"{label}.entity_ids", nonempty=False, ids=True
        )
        unknown_entities = [item for item in entity_ids if item not in entities]
        if unknown_entities:
            raise ValueError(
                f"integrated page {page_id} references unknown entity id(s): "
                + ", ".join(unknown_entities)
            )
        _validate_page_entity_wikilinks(body, entity_ids, page_id)
        result[page_id] = _Page(
            page_id=page_id,
            kind=kind,
            title=title,
            description=description,
            body_markdown=body,
            citations=declared_citations,
            source_ids=source_ids,
            pr_numbers=pr_numbers,
            entity_ids=entity_ids,
        )
    return result


def _parse_coverage(
    records: Sequence[Mapping[str, Any]],
    sources: Mapping[str, _Source],
    pages: Mapping[str, _Page],
    entities: Mapping[str, _Entity],
) -> dict[str, _Coverage]:
    if not records:
        raise ValueError("integrated coverage.jsonl must not be empty")
    result: dict[str, _Coverage] = {}
    document_ids: set[str] = set()
    for index, value in enumerate(records):
        label = f"coverage.jsonl line {index + 1}"
        source_id = _safe_id(value.get("source_id"), f"{label}.source_id")
        if source_id not in sources:
            raise ValueError(f"{label} references unknown source_id: {source_id}")
        if source_id in result:
            raise ValueError(f"duplicate coverage record for source_id: {source_id}")
        document_id = _safe_id(
            value.get("semantic_document_id"), f"{label}.semantic_document_id"
        )
        if document_id != source_id:
            raise ValueError(
                f"coverage semantic_document_id must equal source_id for {source_id}"
            )
        document_key = unicodedata.normalize("NFC", document_id).casefold()
        if document_key in document_ids:
            raise ValueError(f"duplicate coverage semantic_document_id: {document_id}")
        document_ids.add(document_key)
        pr_numbers = _string_tuple(
            value.get("pr_numbers"), f"{label}.pr_numbers", nonempty=True, pr=True
        )
        if pr_numbers != sources[source_id].pr_numbers:
            raise ValueError(f"coverage PR inventory mismatch for source {source_id}")
        page_ids = _string_tuple(
            value.get("page_ids"), f"{label}.page_ids", nonempty=True, ids=True
        )
        entity_ids = _string_tuple(
            value.get("entity_ids"), f"{label}.entity_ids", nonempty=False, ids=True
        )
        if value.get("covered") is not True:
            raise ValueError(f"coverage record must set covered=true for {source_id}")
        unknown_pages = [item for item in page_ids if item not in pages]
        unknown_entities = [item for item in entity_ids if item not in entities]
        if unknown_pages:
            raise ValueError(
                f"coverage record for {source_id} references unknown page id(s): "
                + ", ".join(unknown_pages)
            )
        if unknown_entities:
            raise ValueError(
                f"coverage record for {source_id} references unknown entity id(s): "
                + ", ".join(unknown_entities)
            )
        result[source_id] = _Coverage(
            source_id=source_id,
            semantic_document_id=document_id,
            pr_numbers=pr_numbers,
            page_ids=page_ids,
            entity_ids=entity_ids,
        )
    return result


def _validate_coverage(
    manifest: Mapping[str, Any],
    sources: Mapping[str, _Source],
    pages: Mapping[str, _Page],
    entities: Mapping[str, _Entity],
    coverage: Mapping[str, _Coverage],
) -> None:
    if manifest.get("coverage_complete") is not True:
        raise ValueError("integrated coverage is not complete")
    if set(coverage) != set(sources):
        missing = sorted(set(sources) - set(coverage), key=_stable_key)
        extra = sorted(set(coverage) - set(sources), key=_stable_key)
        details: list[str] = []
        if missing:
            details.append("missing sources: " + ", ".join(missing))
        if extra:
            details.append("unknown sources: " + ", ".join(extra))
        raise ValueError("coverage source set mismatch: " + "; ".join(details))

    page_coverage: dict[str, set[str]] = {page_id: set() for page_id in pages}
    entity_coverage: dict[str, set[str]] = {entity_id: set() for entity_id in entities}
    for source_id, record in coverage.items():
        for page_id in record.page_ids:
            page_coverage[page_id].add(source_id)
        for entity_id in record.entity_ids:
            entity_coverage[entity_id].add(source_id)

    for page_id, page in pages.items():
        declared = set(page.source_ids)
        covered = page_coverage[page_id]
        if not covered:
            raise ValueError(f"integrated page is absent from coverage: {page_id}")
        if declared != covered:
            raise ValueError(
                f"page/source coverage mismatch for {page_id}: "
                f"page={sorted(declared)}, coverage={sorted(covered)}"
            )
    for entity_id, entity in entities.items():
        declared = set(entity.source_ids)
        covered = entity_coverage[entity_id]
        if not covered:
            raise ValueError(f"integrated entity is absent from coverage: {entity_id}")
        if declared != covered:
            raise ValueError(
                f"entity/source coverage mismatch for {entity_id}: "
                f"entity={sorted(declared)}, coverage={sorted(covered)}"
            )


def _validate_citation_evidence(
    citations: Mapping[str, _Citation], sources: Mapping[str, _Source]
) -> dict[str, set[str]]:
    anchors_by_output_stem: dict[str, set[str]] = {}
    anchors_by_slide: dict[tuple[str, int], set[str]] = {}
    for source_id, source in sources.items():
        for slide_number, (_, markdown) in source.slides.items():
            anchors = {
                html_unescape(match.group("id")) for match in _ANCHOR_RE.finditer(markdown)
            }
            anchors_by_slide[(source_id, slide_number)] = anchors
            anchors_by_output_stem[
                f"evidence/{source_id}/slide-{slide_number:04d}"
            ] = anchors | _heading_anchors(markdown)

        # Evidence pages publish every parsed block, including blocks omitted
        # from the semantic document/source-map.  Bind every one of those
        # published blocks back to its content-addressed provenance record so
        # a post-semantic edit to a slide Markdown file cannot be published.
        for local, record in source.provenance_records.items():
            slide_number = int(record["slide_number"])
            element_id = str(record["element_id"])
            slide_entry = source.slides.get(slide_number)
            if slide_entry is None:
                raise ValueError(
                    f"parsed provenance references missing slide: "
                    f"{source_id}/{local}"
                )
            if element_id not in anchors_by_slide[(source_id, slide_number)]:
                raise ValueError(
                    f"parsed provenance references missing evidence anchor: "
                    f"{source_id}/{local}"
                )
            provenance_digest = _digest(
                record.get("content_sha256"),
                f"parsed provenance content SHA-256 for {source_id}/{local}",
            )
            slide_digest = _digest_anchor_content(
                slide_entry[1], local, element_id
            )
            if slide_digest != provenance_digest:
                raise ValueError(
                    f"parsed slide/provenance content SHA-256 mismatch: "
                    f"{source_id}/{local}; expected {provenance_digest}, "
                    f"got {slide_digest}"
                )

    for token, citation in citations.items():
        source = sources[citation.source_id]
        slide_entry = source.slides.get(citation.slide_number)
        if slide_entry is None:
            raise ValueError(
                f"qualified citation references missing slide: {token}"
            )
        anchors = anchors_by_slide[(citation.source_id, citation.slide_number)]
        if citation.element_id not in anchors:
            raise ValueError(
                f"qualified citation references missing evidence anchor: {token}"
            )
        provenance_record = source.provenance_records.get(citation.local)
        if provenance_record is None:
            raise ValueError(f"qualified citation has no parsed provenance: {token}")
        provenance_digest = _digest(
            provenance_record.get("content_sha256"),
            f"parsed provenance content SHA-256 for {token}",
        )
        if provenance_digest != citation.content_sha256:
            raise ValueError(
                f"qualified citation content SHA-256 mismatch: {token}; "
                f"expected {citation.content_sha256}, got {provenance_digest}"
            )
    return anchors_by_output_stem


def _digest_anchor_content(markdown: str, citation: str, element_id: str) -> str:
    anchors = list(_ANCHOR_RE.finditer(markdown))
    target_index = next(
        (
            index
            for index, match in enumerate(anchors)
            if html_unescape(match.group("id")) == element_id
        ),
        None,
    )
    if target_index is None:
        raise ValueError(f"evidence anchor not found for {citation}")
    start = anchors[target_index].end()
    closing_anchor = re.match(r"\s*</a\s*>", markdown[start:], re.IGNORECASE)
    if closing_anchor is None:
        raise ValueError(f"evidence anchor is not explicitly closed for {citation}")
    start += closing_anchor.end()
    end = anchors[target_index + 1].start() if target_index + 1 < len(anchors) else len(markdown)
    end_marker = re.search(
        rf"<!--\s*END\s+(?:BLOCK|TABLE)\s+{re.escape(citation)}\s*-->",
        markdown[start:end],
        re.IGNORECASE,
    )
    if end_marker is not None:
        end = start + end_marker.start()
    value = markdown[start:end]
    value = re.sub(
        rf"(?:\r?\n)+\s*{re.escape(citation)}\s*$", "", value.rstrip()
    ).strip()
    return sha256(value.encode("utf-8")).hexdigest()


def _citation_wikilink(citation: _Citation) -> str:
    target = (
        f"evidence/{citation.source_id}/slide-{citation.slide_number:04d}#"
        + quote(citation.element_id, safe="-._~")
    )
    # Keep the exact source spelling in the evidence label while PR index
    # pages/frontmatter use the source identity's canonical display value.
    pr_label = ", ".join(citation.pr_variants or citation.pr_numbers)
    if not pr_label:
        pr_label = "evidence"
    return f"[[{target}|{pr_label} · slide {citation.slide_number}]]"


def _replace_citations(
    value: str,
    citations: Mapping[str, _Citation],
    links: Mapping[str, str],
    *,
    local_source_id: str | None = None,
    allow_unmapped_local: bool = False,
) -> str:
    _validate_no_malformed_qualified_citations(value, "Markdown")

    def replace_qualified(match: re.Match[str]) -> str:
        token = match.group(0)
        if token not in citations:
            raise ValueError(f"Markdown contains unknown qualified citation: {token}")
        return links[token]

    result = _QUALIFIED_CITATION_RE.sub(replace_qualified, value)
    if local_source_id is None:
        return result
    local_map = {
        citation.local: citation
        for citation in citations.values()
        if citation.source_id == local_source_id
    }

    def replace_local(match: re.Match[str]) -> str:
        token = match.group(0)
        citation = local_map.get(token)
        if citation is not None:
            return links[citation.qualified]
        if not allow_unmapped_local:
            raise ValueError(
                f"semantic Markdown contains unmapped local citation for "
                f"{local_source_id}: {token}"
            )
        slide_number = int(match.group("slide"))
        element_id = match.group("element")
        target = (
            f"evidence/{local_source_id}/slide-{slide_number:04d}#"
            + quote(element_id, safe="-._~")
        )
        return f"[[{target}|slide {slide_number}]]"

    result = _LOCAL_CITATION_RE.sub(replace_local, result)
    if "[slide-" in result:
        position = result.index("[slide-")
        sample = result[position : position + 120].splitlines()[0]
        raise ValueError(f"Markdown contains malformed local citation: {sample!r}")
    return result


def _prepare_input_markdown(
    value: str, *, label: str, allow_frontmatter: bool
) -> str:
    _validate_text_controls(value, label)
    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = text.removeprefix("\ufeff")
    stripped = text.lstrip("\n")
    if stripped.startswith("---\n"):
        if not allow_frontmatter:
            raise ValueError(f"{label} must not contain publisher-owned frontmatter")
        offset = len(text) - len(stripped)
        closing = text.find("\n---\n", offset + 4)
        if closing < 0:
            raise ValueError(f"{label} has unterminated frontmatter")
        text = text[closing + len("\n---\n") :]
    _validate_no_dangerous_html(text, label)
    if "file://" in text.casefold():
        raise ValueError(f"{label} contains a forbidden file:// URL")
    return text.strip()


def _rewrite_local_links(
    markdown: str,
    *,
    input_path: Path,
    output_path: str,
    source: _Source,
    assets: _AssetStore,
) -> str:
    def inline(match: re.Match[str]) -> str:
        image = bool(match.group("image"))
        label = match.group("label")
        raw_target = match.group("target")
        target, _ = _split_markdown_target(raw_target)
        rewritten = _rewrite_link_target(
            target,
            is_image=image,
            input_path=input_path,
            output_path=output_path,
            source=source,
            assets=assets,
        )
        marker = "!" if image else ""
        return f"{marker}[{label}]({rewritten})"

    value = _INLINE_LINK_RE.sub(inline, markdown)

    def reference(match: re.Match[str]) -> str:
        raw_target = match.group("target")
        target = raw_target[1:-1] if raw_target.startswith("<") else raw_target
        rewritten = _rewrite_link_target(
            target,
            is_image=False,
            input_path=input_path,
            output_path=output_path,
            source=source,
            assets=assets,
        )
        return f"{match.group('prefix')}<{rewritten}>{match.group('rest')}"

    value = _REFERENCE_LINK_RE.sub(reference, value)

    def html_link(match: re.Match[str]) -> str:
        attribute = match.group("attribute").casefold()
        rewritten = _rewrite_link_target(
            match.group("target"),
            is_image=attribute == "src",
            input_path=input_path,
            output_path=output_path,
            source=source,
            assets=assets,
        )
        escaped = html_escape(rewritten, quote=True)
        quote_character = match.group("quote")
        return f"{match.group('prefix')}{quote_character}{escaped}{quote_character}"

    value = _HTML_LINK_RE.sub(html_link, value)
    return value


def _split_markdown_target(value: str) -> tuple[str, str]:
    stripped = value.strip()
    if not stripped:
        raise ValueError("Markdown link target cannot be empty")
    if stripped.startswith("<"):
        closing = stripped.find(">")
        if closing < 0:
            raise ValueError(f"unterminated angle-bracket Markdown target: {value!r}")
        return stripped[1:closing], stripped[closing + 1 :].strip()
    # Titles are intentionally discarded.  First try the complete target so a
    # legacy absolute path containing spaces can still be safely copied.
    return stripped, ""


def _rewrite_link_target(
    target: str,
    *,
    is_image: bool,
    input_path: Path,
    output_path: str,
    source: _Source,
    assets: _AssetStore,
) -> str:
    target = html_unescape(target).strip()
    if not target:
        raise ValueError("Markdown/HTML link target cannot be empty")
    if any(character in target for character in ("\x00", "\r", "\n")):
        raise ValueError("Markdown/HTML link target contains control characters")
    if target.startswith("#"):
        if is_image:
            raise ValueError("image target cannot be a local fragment")
        return target
    windows_absolute = PureWindowsPath(target).is_absolute() or bool(
        re.match(r"^[A-Za-z]:[\\/]", target)
    )
    parsed = urlsplit(target)
    scheme = parsed.scheme.casefold()
    if scheme and not windows_absolute:
        if scheme == "data" and is_image:
            if not target.casefold().startswith("data:image/"):
                raise ValueError("only data:image assets are allowed")
            return target
        if scheme in {"https", "http", "mailto"} and not is_image:
            return target
        raise ValueError(f"forbidden or non-self-contained link scheme: {scheme}")
    if target.startswith(("//", "\\\\")):
        raise ValueError(f"UNC/network asset targets are forbidden: {target!r}")

    path_part, separator, fragment = target.partition("#")
    path_part = unquote(path_part)
    source_root = source.root.resolve(strict=True)
    candidates: list[Path] = []
    raw_path = Path(path_part)
    if windows_absolute or raw_path.is_absolute():
        candidates.append(raw_path)
    else:
        candidates.extend(
            (
                input_path.parent / raw_path,
                source.root / raw_path,
                source.root / "parsed" / raw_path,
            )
        )
    resolved: Path | None = None
    for candidate in candidates:
        try:
            current = candidate.resolve(strict=True)
        except (FileNotFoundError, OSError):
            continue
        if current.is_relative_to(source_root):
            resolved = current
            break
    if resolved is None:
        raise ValueError(
            f"local link target is missing or escapes source {source.source_id}: {target!r}"
        )
    if resolved.is_symlink() or not resolved.is_file():
        raise ValueError(f"local link target is not a regular file: {target!r}")

    slides_root = (source.root / "parsed" / "corpus" / "slides").resolve(strict=True)
    if resolved.is_relative_to(slides_root) and resolved.suffix.casefold() == ".md":
        match = re.fullmatch(r"slide-(?P<number>\d+)\.md", resolved.name)
        if match is None:
            raise ValueError(f"unsafe local slide link: {target!r}")
        if int(match.group("number")) <= 0:
            raise ValueError(f"local slide link has a non-positive number: {target!r}")
        destination_target = (
            f"content/evidence/{source.source_id}/"
            f"slide-{int(match.group('number')):04d}.md"
        )
        relative = _relative_output_link(output_path, destination_target)
        if separator:
            relative += "#" + quote(unquote(fragment), safe="-._~")
        return relative
    if resolved.suffix.casefold() == ".md":
        raise ValueError(f"unsupported local Markdown link: {target!r}")

    asset_output = assets.add(source, resolved)
    relative = _relative_output_link(output_path, asset_output)
    if separator:
        relative += "#" + quote(unquote(fragment), safe="-._~")
    return relative


def _relative_output_link(source_path: str, target_path: str) -> str:
    source_parent = PurePosixPath(source_path).parent
    relative = os.path.relpath(
        str(PurePosixPath(target_path)), start=str(source_parent)
    ).replace("\\", "/")
    return quote(relative, safe="/.-_~")


def _render_markdown_page(
    *,
    title: str,
    description: str,
    aliases: Sequence[str],
    tags: Sequence[str],
    pr_numbers: Sequence[str],
    source_ids: Sequence[str],
    body: str,
) -> str:
    title_value = _one_line(title, "Quartz title")
    description_value = _frontmatter_description(description)
    alias_values = tuple(_one_line(item, "Quartz alias") for item in aliases)
    tag_values = tuple(_one_line(item, "Quartz tag") for item in tags)
    pr_values = tuple(_pr_number(item, "Quartz PR number") for item in pr_numbers)
    source_values = tuple(_safe_id(item, "Quartz source id") for item in source_ids)
    lines = [
        "---",
        f"title: {_yaml_scalar(title_value)}",
        f"description: {_yaml_scalar(description_value)}",
        f"aliases: {_yaml_array(alias_values)}",
        f"tags: {_yaml_array(tag_values)}",
        f"pr_numbers: {_yaml_array(pr_values)}",
        f"source_ids: {_yaml_array(source_values)}",
        "draft: false",
        "---",
        "",
        body.strip(),
        "",
    ]
    value = "\n".join(lines)
    _validate_generated_frontmatter(value)
    return value


def _with_heading(title: str, body: str) -> str:
    clean_title = _escape_heading(_one_line(title, "Markdown heading"))
    value = body.strip()
    if value:
        first = value.splitlines()[0].strip()
        if _HEADING_RE.fullmatch(first):
            value = "\n".join(value.splitlines()[1:]).lstrip()
    return f"# {clean_title}" + (f"\n\n{value}" if value else "")


def _render_entity_body(
    entity: _Entity, citation_links: Mapping[str, str]
) -> str:
    lines = [f"# {_escape_heading(entity.canonical_name)}", ""]
    description = entity.description
    if description.strip():
        _validate_body_citations(description, entity.citations, f"entity {entity.entity_id}")
        rendered = description
        for token in entity.citations:
            rendered = rendered.replace(token, citation_links[token])
        if not _QUALIFIED_CITATION_RE.search(description):
            rendered = rendered.rstrip() + " " + " ".join(
                citation_links[token] for token in entity.citations
            )
        lines.extend((rendered.strip(), ""))
    if entity.aliases:
        lines.extend(("## Aliases", ""))
        lines.extend(f"- {_escape_markdown_text(alias)}" for alias in entity.aliases)
        lines.append("")
    lines.extend(("## Evidence", ""))
    lines.extend(f"- {citation_links[token]}" for token in entity.citations)
    return "\n".join(lines).rstrip()


def _append_entity_links(
    body: str, entity_ids: Sequence[str], entities: Mapping[str, _Entity]
) -> str:
    if not entity_ids:
        return body.strip()
    links = [f"- [[entities/{entity_id}]]" for entity_id in entity_ids]
    return body.rstrip() + "\n\n## Related entities\n\n" + "\n".join(links)


def _render_pr_body(
    pr_number: str,
    sources: Mapping[str, _Source],
    pages: Mapping[str, _Page],
    entities: Mapping[str, _Entity],
) -> str:
    lines = [f"# {_escape_heading(pr_number)}", "", "## Sources", ""]
    matching_sources = [
        source_id
        for source_id in sorted(sources, key=_stable_key)
        if pr_number in sources[source_id].pr_numbers
    ]
    lines.extend(f"- [[sources/{source_id}]]" for source_id in matching_sources)
    matching_pages = [
        page_id
        for page_id in sorted(pages, key=_stable_key)
        if pr_number in pages[page_id].pr_numbers
    ]
    if matching_pages:
        lines.extend(("", "## Topics", ""))
        lines.extend(f"- [[topics/{page_id}]]" for page_id in matching_pages)
    matching_entities = [
        entity_id
        for entity_id in sorted(entities, key=_stable_key)
        if pr_number in entities[entity_id].pr_numbers
    ]
    if matching_entities:
        lines.extend(("", "## Entities", ""))
        lines.extend(f"- [[entities/{entity_id}]]" for entity_id in matching_entities)
    return "\n".join(lines).rstrip()


def _render_index_body(
    site_title: str,
    sources: Mapping[str, _Source],
    pages: Mapping[str, _Page],
    entities: Mapping[str, _Entity],
    pr_slugs: Mapping[str, str],
) -> str:
    lines = [f"# {_escape_heading(site_title)}", ""]
    if pages:
        lines.extend(("## Topics", ""))
        lines.extend(
            f"- [[topics/{page_id}]]"
            for page_id in sorted(pages, key=_stable_key)
        )
    if entities:
        lines.extend(("", "## Entities", ""))
        lines.extend(
            f"- [[entities/{entity_id}]]"
            for entity_id in sorted(entities, key=_stable_key)
        )
    lines.extend(("", "## PR inventory", ""))
    lines.extend(
        f"- [[prs/{pr_slugs[pr_number]}]]"
        for pr_number in sorted(pr_slugs, key=_stable_key)
    )
    lines.extend(("", "## Sources", ""))
    lines.extend(
        f"- [[sources/{source_id}]]"
        for source_id in sorted(sources, key=_stable_key)
    )
    return "\n".join(lines).rstrip()


def _render_readme(site_title: str) -> str:
    return (
        f"# {_escape_heading(site_title)}\n\n"
        "This directory is a deterministic, self-contained Quartz input.\n\n"
        "- Put or link `content/` at the Quartz project content root.\n"
        "- `quartz-manifest.json` records every published file hash.\n"
        "- Topic and entity claims link to source-faithful evidence pages.\n"
    )


def _render_quartz_manifest(
    *,
    site_title: str,
    integrated_manifest_sha256: str,
    sources: Mapping[str, _Source],
    entities: Mapping[str, _Entity],
    pages: Mapping[str, _Page],
    pr_slugs: Mapping[str, str],
    files: Mapping[str, bytes],
) -> str:
    file_records = {
        path: {"sha256": sha256(raw).hexdigest(), "size": len(raw)}
        for path, raw in sorted(files.items(), key=lambda item: _stable_key(item[0]))
    }
    content_page_count = sum(
        path.startswith("content/") and path.endswith(".md") for path in files
    )
    assets = sorted(
        (path for path in files if path.startswith("content/assets/")),
        key=_stable_key,
    )
    value = {
        "schema_version": QUARTZ_SCHEMA_VERSION,
        "site_title": site_title,
        "source_integrated_manifest_sha256": integrated_manifest_sha256,
        "source_count": len(sources),
        "topic_count": len(pages),
        "entity_count": len(entities),
        "pr_count": len(pr_slugs),
        "page_count": content_page_count,
        "asset_count": len(assets),
        "sources": [
            {
                "source_id": source_id,
                "pr_numbers": list(sources[source_id].pr_numbers),
                "source_sha256": sources[source_id].source_sha256,
            }
            for source_id in sorted(sources, key=_stable_key)
        ],
        "prs": [
            {"pr_number": pr_number, "page": f"content/prs/{pr_slugs[pr_number]}.md"}
            for pr_number in sorted(pr_slugs, key=_stable_key)
        ],
        "files": file_records,
    }
    return json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2) + "\n"


def _validate_rendered_tree(
    files: Mapping[str, bytes], evidence_anchors: Mapping[str, set[str]]
) -> None:
    markdown: dict[str, str] = {}
    for path, raw in files.items():
        if not path.endswith(".md"):
            continue
        try:
            value = raw.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise ValueError(f"rendered Markdown is not UTF-8 for {path}: {exc}") from exc
        if path.startswith("content/"):
            _validate_generated_frontmatter(value)
        _validate_no_dangerous_html(value, path)
        _validate_no_malformed_qualified_citations(value, path)
        if _QUALIFIED_CITATION_RE.search(value) or _LOCAL_CITATION_RE.search(value):
            raise ValueError(f"unmaterialized citation remains in Quartz output: {path}")
        if "[slide-" in value or "[@" in value:
            raise ValueError(f"malformed/unmaterialized citation remains in Quartz output: {path}")
        markdown[path] = value

    page_stems = {
        path[len("content/") : -len(".md")]: path
        for path in markdown
        if path.startswith("content/")
    }
    stem_keys: dict[str, str] = {}
    for stem in page_stems:
        key = unicodedata.normalize("NFC", stem).casefold()
        previous = stem_keys.get(key)
        if previous is not None and previous != stem:
            raise ValueError(f"Quartz page casefold collision: {previous!r}, {stem!r}")
        stem_keys[key] = stem

    anchors: dict[str, set[str]] = {}
    for stem, path in page_stems.items():
        anchors[stem] = _markdown_anchors(markdown[path])
    for stem, explicit in evidence_anchors.items():
        anchors.setdefault(stem, set()).update(explicit)

    for path, value in markdown.items():
        if not path.startswith("content/"):
            continue
        for match in _WIKILINK_RE.finditer(value):
            raw_target = match.group("value")
            target_part = raw_target.split("|", 1)[0].strip()
            target_path, marker, fragment = target_part.partition("#")
            target_path = target_path.removesuffix(".md").lstrip("/")
            target_path = _safe_relative_posix(target_path, f"wikilink in {path}")
            target_key = unicodedata.normalize("NFC", target_path).casefold()
            canonical = stem_keys.get(target_key)
            if canonical is None:
                raise ValueError(f"unresolved Quartz wikilink in {path}: {raw_target}")
            if marker:
                decoded_fragment = unquote(fragment)
                if decoded_fragment not in anchors.get(canonical, set()):
                    raise ValueError(
                        f"unresolved Quartz wikilink fragment in {path}: {raw_target}"
                    )
        _validate_standard_links(path, value, files, anchors)


def _validate_standard_links(
    markdown_path: str,
    value: str,
    files: Mapping[str, bytes],
    anchors: Mapping[str, set[str]],
) -> None:
    targets = [match.group("target") for match in _INLINE_LINK_RE.finditer(value)]
    targets.extend(match.group("target") for match in _REFERENCE_LINK_RE.finditer(value))
    targets.extend(match.group("target") for match in _HTML_LINK_RE.finditer(value))
    for raw_target in targets:
        target, _ = _split_markdown_target(raw_target)
        target = html_unescape(target)
        if target.startswith("#"):
            current_stem = markdown_path[len("content/") : -len(".md")]
            if unquote(target[1:]) not in anchors.get(current_stem, set()):
                raise ValueError(
                    f"unresolved local Markdown fragment in {markdown_path}: {target}"
                )
            continue
        parsed = urlsplit(target)
        if parsed.scheme.casefold() in {"http", "https", "mailto", "data"}:
            continue
        if parsed.scheme or PureWindowsPath(target).is_absolute() or target.startswith("/"):
            raise ValueError(f"absolute/unsafe link in {markdown_path}: {target}")
        path_value = unquote(parsed.path)
        resolved = _normalise_relative_output_path(
            str(PurePosixPath(markdown_path).parent), path_value
        )
        if resolved not in files:
            raise ValueError(f"unresolved Markdown link in {markdown_path}: {target}")
        if parsed.fragment and resolved.endswith(".md") and resolved.startswith("content/"):
            stem = resolved[len("content/") : -len(".md")]
            if unquote(parsed.fragment) not in anchors.get(stem, set()):
                raise ValueError(
                    f"unresolved Markdown fragment in {markdown_path}: {target}"
                )


def _commit_files(destination: Path, files: Mapping[str, bytes]) -> None:
    _ensure_empty_destination(destination)
    staging_value = tempfile.mkdtemp(
        prefix=f".{destination.name or 'quartz'}.staging-", dir=destination.parent
    )
    staging = Path(staging_value).resolve()
    expected_parent = destination.parent.resolve(strict=True)
    if staging.parent != expected_parent:
        shutil.rmtree(staging, ignore_errors=True)
        raise RuntimeError("Quartz staging directory was created on the wrong filesystem")
    committed = False
    removed_empty_destination = False
    try:
        for relative in _OUTPUT_DIRECTORIES:
            (staging / PurePosixPath(relative)).mkdir(parents=True, exist_ok=True)
        for relative, raw in sorted(files.items(), key=lambda item: _stable_key(item[0])):
            target = staging / PurePosixPath(relative)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(raw)
        _ensure_empty_destination(destination)
        if destination.exists():
            destination.rmdir()
            removed_empty_destination = True
        os.replace(staging, destination)
        committed = True
    except Exception:
        if removed_empty_destination and not destination.exists():
            try:
                destination.mkdir()
            except OSError:
                pass
        raise
    finally:
        if not committed and staging.exists() and staging.parent == expected_parent:
            shutil.rmtree(staging)


def _validate_declared_citations(
    declared: Sequence[str],
    citations: Mapping[str, _Citation],
    label: str,
) -> None:
    unknown = [token for token in declared if token not in citations]
    if unknown:
        raise ValueError(
            f"{label} declares unknown qualified citation(s): " + ", ".join(unknown)
        )
    malformed = [
        token for token in declared if _QUALIFIED_CITATION_RE.fullmatch(token) is None
    ]
    if malformed:
        raise ValueError(
            f"{label} declares malformed qualified citation(s): "
            + ", ".join(malformed)
        )


def _validate_body_citations(
    body: str, declared: Sequence[str], label: str
) -> None:
    _validate_no_malformed_qualified_citations(body, label)
    embedded = {match.group(0) for match in _QUALIFIED_CITATION_RE.finditer(body)}
    undeclared = sorted(embedded - set(declared), key=_stable_key)
    if undeclared:
        raise ValueError(
            f"{label} body has undeclared citation(s): " + ", ".join(undeclared)
        )


def _validate_no_malformed_qualified_citations(value: str, label: str) -> None:
    scrubbed = _QUALIFIED_CITATION_RE.sub("", value)
    if "[@" in scrubbed:
        position = scrubbed.index("[@")
        sample = scrubbed[position : position + 120].splitlines()[0]
        raise ValueError(f"{label} contains malformed qualified citation: {sample!r}")


def _validate_lineage_lists(
    declared: Sequence[str],
    source_ids: Sequence[str],
    pr_numbers: Sequence[str],
    citations: Mapping[str, _Citation],
    label: str,
    sources: Mapping[str, _Source],
) -> None:
    derived_sources = {
        citations[token].source_id for token in declared if token in citations
    }
    # Integration uses citation-local PRs when that source has at least one
    # PR-bearing citation in the record, and otherwise falls back to the
    # source's complete PR inventory.  Recompute that rule here instead of
    # trusting the redundant arrays emitted by integration.
    derived_prs: set[str] = set()
    for source_id in derived_sources:
        specific = {
            pr
            for token in declared
            if token in citations and citations[token].source_id == source_id
            for pr in citations[token].pr_numbers
        }
        derived_prs.update(specific or sources[source_id].pr_numbers)
    if set(source_ids) != derived_sources:
        raise ValueError(
            f"{label} source_ids do not match citation lineage: "
            f"declared={sorted(source_ids)}, derived={sorted(derived_sources)}"
        )
    if set(pr_numbers) != derived_prs:
        raise ValueError(
            f"{label} pr_numbers do not match citation lineage: "
            f"declared={sorted(pr_numbers)}, derived={sorted(derived_prs)}"
        )


def _numeric_tokens(value: str) -> set[str]:
    cleaned = _QUALIFIED_CITATION_RE.sub("", unicodedata.normalize("NFKC", value))
    cleaned = re.sub(r"\]\([^\n)]*\)", "]", cleaned)
    cleaned = re.sub(r"(?m)^\s*\d+[.)]\s+", "", cleaned)
    return {_normal_number(match.group(0)) for match in _NUMBER_RE.finditer(cleaned)}


def _normal_number(value: str) -> str:
    percent = value.endswith("%")
    cleaned = value.rstrip("%").replace(",", "")
    cleaned = cleaned.removeprefix("+")
    if "." in cleaned:
        cleaned = cleaned.rstrip("0").rstrip(".")
    if cleaned in {"", "-0"}:
        cleaned = "0"
    return cleaned + ("%" if percent else "")


def _identifier_tokens(value: str) -> tuple[str, ...]:
    """Mirror the integrated artifact's case-insensitive token inventory."""

    values: list[str] = []
    seen: set[str] = set()
    for match in _IDENTIFIER_TOKEN_RE.finditer(value):
        token = unicodedata.normalize("NFC", match.group(0))
        key = token.casefold()
        if key not in seen:
            seen.add(key)
            values.append(token)
    return tuple(values)


def _validate_page_entity_wikilinks(
    body: str, entity_ids: Sequence[str], page_id: str
) -> None:
    declared = set(entity_ids)
    for match in _WIKILINK_RE.finditer(body):
        raw = match.group("value").split("|", 1)[0].split("#", 1)[0]
        target = raw.removesuffix(".md").strip().lstrip("/")
        if target.startswith("entities/"):
            entity_id = target[len("entities/") :]
            if entity_id not in declared:
                raise ValueError(
                    f"integrated page {page_id} links undeclared entity: {entity_id}"
                )


def _assign_pr_slugs(values: Iterable[str]) -> dict[str, str]:
    prs = sorted(set(values), key=_stable_key)
    bases: dict[str, list[str]] = {}
    for value in prs:
        base = _slugify(value)
        bases.setdefault(unicodedata.normalize("NFC", base).casefold(), []).append(value)
    result: dict[str, str] = {}
    used: set[str] = set()
    for value in prs:
        base = _slugify(value)
        collision_group = bases[unicodedata.normalize("NFC", base).casefold()]
        candidate = (
            base
            if len(collision_group) == 1
            else f"{base}--{sha256(value.encode('utf-8')).hexdigest()[:10]}"
        )
        key = unicodedata.normalize("NFC", candidate).casefold()
        if key in used:
            candidate = f"{base}--{sha256(value.encode('utf-8')).hexdigest()}"
            key = unicodedata.normalize("NFC", candidate).casefold()
        if key in used:
            raise ValueError(f"unable to assign unique PR page path for {value!r}")
        used.add(key)
        result[value] = candidate
    return result


def _slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value).casefold()
    normalized = re.sub(r"[^\w.-]+", "-", normalized, flags=re.UNICODE)
    normalized = re.sub(r"-+", "-", normalized).strip(" .-")
    if not normalized:
        normalized = "pr"
    if normalized.split(".", 1)[0].casefold() in _WINDOWS_RESERVED:
        normalized = "_" + normalized
    return normalized[:80].rstrip(" .") or "pr"


def _tag_token(value: str) -> str:
    token = _slugify(value).replace(".", "-")
    return token or "item"


def _frontmatter_description(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("Quartz description must be a string")
    value = _QUALIFIED_CITATION_RE.sub("", value)
    value = _LOCAL_CITATION_RE.sub("", value)
    return " ".join(value.split()).strip()[:500]


def _yaml_scalar(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def _yaml_array(values: Sequence[str]) -> str:
    return json.dumps(list(values), ensure_ascii=False)


def _validate_generated_frontmatter(value: str) -> None:
    if not value.startswith("---\n"):
        raise ValueError("Quartz Markdown is missing leading frontmatter")
    closing = value.find("\n---\n", 4)
    if closing < 0:
        raise ValueError("Quartz Markdown has unterminated frontmatter")
    lines = value[4:closing].splitlines()
    expected = {
        "title",
        "description",
        "aliases",
        "tags",
        "pr_numbers",
        "source_ids",
        "draft",
    }
    seen: set[str] = set()
    parsed: dict[str, Any] = {}
    for line in lines:
        key, separator, raw = line.partition(":")
        if not separator or not key:
            raise ValueError(f"invalid generated frontmatter line: {line!r}")
        if key in seen:
            raise ValueError(f"duplicate generated frontmatter key: {key}")
        seen.add(key)
        raw = raw.strip()
        if key == "draft":
            if raw != "false":
                raise ValueError("generated Quartz page must set draft: false")
            parsed[key] = False
        else:
            try:
                parsed[key] = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise ValueError(f"invalid generated frontmatter value for {key}") from exc
    if seen != expected:
        raise ValueError(
            f"generated frontmatter keys mismatch: expected {sorted(expected)}, got {sorted(seen)}"
        )
    if not isinstance(parsed["title"], str) or not parsed["title"]:
        raise ValueError("generated frontmatter title must be a non-empty string")
    if not isinstance(parsed["description"], str):
        raise TypeError("generated frontmatter description must be a string")
    for key in ("aliases", "tags", "pr_numbers", "source_ids"):
        if not isinstance(parsed[key], list) or not all(
            isinstance(item, str) for item in parsed[key]
        ):
            raise ValueError(f"generated frontmatter {key} must be a string array")


def _markdown_anchors(value: str) -> set[str]:
    anchors = {
        html_unescape(match.group("id")) for match in _ANCHOR_RE.finditer(value)
    }
    anchors.update(_heading_anchors(value))
    return anchors


def _heading_anchors(value: str) -> set[str]:
    anchors: set[str] = set()
    counts: dict[str, int] = {}
    in_fence = False
    fence_marker = ""
    for line in value.splitlines():
        stripped = line.lstrip()
        if stripped.startswith(("```", "~~~")):
            marker = stripped[:3]
            if not in_fence:
                in_fence = True
                fence_marker = marker
            elif marker == fence_marker:
                in_fence = False
            continue
        if in_fence:
            continue
        match = _HEADING_RE.fullmatch(line.strip())
        if match is None:
            continue
        slug = _heading_slug(match.group("title"))
        number = counts.get(slug, 0)
        counts[slug] = number + 1
        anchors.add(slug if number == 0 else f"{slug}-{number}")
    return anchors


def _heading_slug(value: str) -> str:
    value = re.sub(r"[`*_~\[\]()]", "", value).strip().casefold()
    value = re.sub(r"[^\w\-\s]", "", value, flags=re.UNICODE)
    return re.sub(r"[\s-]+", "-", value).strip("-")


def _normalise_relative_output_path(base: str, target: str) -> str:
    combined = PurePosixPath(base) / PurePosixPath(target)
    parts: list[str] = []
    for part in combined.parts:
        if part in {"", "."}:
            continue
        if part == "..":
            if not parts:
                raise ValueError(f"link escapes Quartz output root: {target!r}")
            parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def _safe_relative_posix(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{label} must be a non-empty string")
    if "\\" in value or "\x00" in value or ":" in value:
        raise ValueError(f"{label} is not a safe relative POSIX path: {value!r}")
    path = PurePosixPath(value)
    if path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"{label} is not a safe relative POSIX path: {value!r}")
    for part in path.parts:
        if part.endswith((" ", ".")):
            raise ValueError(f"{label} has a Windows-unsafe path segment: {part!r}")
        if part.split(".", 1)[0].casefold() in _WINDOWS_RESERVED:
            raise ValueError(f"{label} uses a Windows-reserved path segment: {part!r}")
    return path.as_posix()


def _safe_id(value: Any, label: str) -> str:
    if not isinstance(value, str) or _SAFE_ID_RE.fullmatch(value) is None:
        raise ValueError(f"{label} must be a safe ASCII identifier")
    if value.endswith(".") or value.split(".", 1)[0].casefold() in _WINDOWS_RESERVED:
        raise ValueError(f"{label} is Windows-reserved or unsafe: {value!r}")
    return value


def _element_id(value: Any, label: str) -> str:
    text = _one_line(value, label)
    if any(character.isspace() for character in text) or any(
        character in text for character in "[]#"
    ):
        raise ValueError(f"{label} is not a valid citation element id")
    return text


def _one_line(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{label} must be a string")
    if value != value.strip() or not value or "\r" in value or "\n" in value:
        raise ValueError(f"{label} must be a non-empty, trimmed, one-line string")
    _validate_text_controls(value, label)
    return unicodedata.normalize("NFC", value)


def _text(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{label} must be a string")
    _validate_text_controls(value, label)
    return unicodedata.normalize(
        "NFC", value.replace("\r\n", "\n").replace("\r", "\n")
    )


def _string_tuple(
    value: Any,
    label: str,
    *,
    nonempty: bool,
    ids: bool = False,
    pr: bool = False,
) -> tuple[str, ...]:
    if not isinstance(value, list) or (nonempty and not value):
        qualifier = "non-empty " if nonempty else ""
        raise ValueError(f"{label} must be a {qualifier}array of strings")
    values: list[str] = []
    seen: set[str] = set()
    for index, item in enumerate(value):
        if ids:
            parsed = _safe_id(item, f"{label}[{index}]")
        elif pr:
            parsed = _pr_number(item, f"{label}[{index}]")
        else:
            parsed = _one_line(item, f"{label}[{index}]")
        key = unicodedata.normalize("NFC", parsed).casefold()
        if key in seen:
            raise ValueError(f"{label} contains a duplicate value: {parsed!r}")
        seen.add(key)
        values.append(parsed)
    return tuple(values)


def _identifier_token_tuple(value: Any, label: str) -> tuple[str, ...]:
    """Read a legacy identifier inventory and collapse case-only duplicates.

    Identifier inventories are set-like.  Older integrated artifacts emitted
    both ``LOT NO.`` and ``Lot No.`` while the publisher already treated them
    as the same value.  Preserve the first spelling so Quartz can be resumed
    without weakening duplicate validation for any other artifact field.
    """

    if not isinstance(value, list):
        raise ValueError(f"{label} must be an array of strings")
    values: list[str] = []
    seen: set[str] = set()
    for index, item in enumerate(value):
        parsed = _one_line(item, f"{label}[{index}]")
        key = unicodedata.normalize("NFC", parsed).casefold()
        if key not in seen:
            seen.add(key)
            values.append(parsed)
    return tuple(values)


def _pr_number(value: Any, label: str) -> str:
    text = _one_line(value, label)
    if len(text) > 128 or any(character in text for character in "[]|#"):
        raise ValueError(f"{label} contains unsafe PR identifier characters")
    return text


def _digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or _DIGEST_RE.fullmatch(value) is None:
        raise ValueError(f"{label} must be a 64-character SHA-256 hex digest")
    return value.casefold()


def _non_negative_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"{label} must be a non-negative integer")
    return value


def _positive_int(value: Any, label: str) -> int:
    result = _non_negative_int(value, label)
    if result == 0:
        raise ValueError(f"{label} must be positive")
    return result


def _validate_text_controls(value: str, label: str) -> None:
    for character in value:
        codepoint = ord(character)
        if codepoint == 0 or (codepoint < 32 and character not in "\t\r\n"):
            raise ValueError(f"{label} contains a forbidden control character")


def _validate_no_dangerous_html(value: str, label: str) -> None:
    if re.search(r"<\s*/?\s*(?:script|iframe|object|embed)\b", value, re.IGNORECASE):
        raise ValueError(f"{label} contains unsafe active HTML")
    if re.search(r"\bon[a-z]+\s*=", value, re.IGNORECASE):
        raise ValueError(f"{label} contains an unsafe HTML event handler")
    if re.search(r"(?:href|src)\s*=\s*[\"']?\s*javascript:", value, re.IGNORECASE):
        raise ValueError(f"{label} contains a javascript: URL")


def _escape_heading(value: str) -> str:
    return value.replace("\\", "\\\\").replace("#", "\\#")


def _escape_markdown_text(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
        .replace("[", "\\[")
        .replace("]", "\\]")
        .replace("*", "\\*")
        .replace("_", "\\_")
    )


def _stable_key(value: Any) -> tuple[str, str]:
    text = str(value)
    return (unicodedata.normalize("NFC", text).casefold(), text)


__all__ = [
    "INTEGRATED_SCHEMA_VERSION",
    "QUARTZ_SCHEMA_VERSION",
    "QuartzExport",
    "publish_quartz",
]
