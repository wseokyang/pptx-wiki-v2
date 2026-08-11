"""Per-presentation semantic dossiers for the collection pipeline.

This module intentionally sits beside :mod:`pptx_wiki.semantic` instead of
changing its single-presentation artifact.  A collection dossier has one
canonical Markdown document per unique PPTX and adds two collection-specific
contracts:

* PR identifiers are extracted from immutable provenance and injected by
  code, never entrusted to the language model; and
* every source block receives an auditable keep/omit/duplicate decision.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from hashlib import sha256
import json
from pathlib import Path
import re
from typing import Any, Callable, Mapping, Sequence
import unicodedata

from lxml import html as lxml_html
import yaml

from .semantic import SemanticConfig
from .synthesis import (
    CITATION_RE,
    ChatBackend,
    Topic,
    _discover_topics,
    _generate_topic_page,
    _numeric_tokens,
    _slide_topics,
    _verbatim_evidence_page,
    _without_top_heading,
)
from .wiki_output import load_provenance


SOURCE_SEMANTIC_SCHEMA_VERSION = "pptx-wiki.source-semantic.v1"

# The labelled form accepts identifiers such as ``123456``, ``PR-123456`` or
# ``A-123456-R2``.  The unlabelled form deliberately requires the PR prefix to
# avoid treating every number in an analysis deck as an identifier.
_PR_LABEL_PATTERN = r"(?:PR[ \t]*(?:번호|NO\.?|NUMBER|#)|의뢰[ \t]*번호)"
_LABELLED_PR_RE = re.compile(
    rf"(?i)(?<![A-Z0-9]){_PR_LABEL_PATTERN}[ \t]*[:：#=-]?[ \t]*"
    r"(?P<value>(?:PR[ \t_-]*)?[A-Z0-9][A-Z0-9._/\-]{0,63})"
)
_PREFIXED_PR_RE = re.compile(
    r"(?i)(?<![A-Z0-9])(?P<value>PR[ \t_-]+(?:[A-Z]+[ \t_-]+)?\d[A-Z0-9._/\-]{0,63})"
    r"(?![A-Z0-9])"
)
_PR_LIKE_TOKEN_RE = re.compile(
    r"(?i)(?<![A-Z0-9])(?P<value>[A-Z]{2,6}[ \t_-]+[A-Z0-9][A-Z0-9._/\-]{0,63})"
    r"(?![A-Z0-9])"
)
_PR_LABEL_RE = re.compile(rf"(?i)(?<![A-Z0-9]){_PR_LABEL_PATTERN}")
_BARE_PR_VALUE_RE = re.compile(r"(?i)(?:PR[ \t_-]*)?[A-Z0-9][A-Z0-9._/\-]{0,63}")
_DIGEST_RE = re.compile(r"^[0-9a-f]{64}$")


@dataclass(frozen=True, slots=True)
class SourceIdentity:
    source_id: str
    source_name: str
    source_sha256: str
    pr_numbers: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class SourceSemanticExport:
    output_dir: Path
    manifest_path: Path
    documents_path: Path
    decisions_path: Path
    markdown_path: Path
    source_id: str
    pr_numbers: tuple[str, ...]
    selected_citations: tuple[str, ...]
    omitted_citations: tuple[str, ...]
    fallback: bool
    warnings: tuple[str, ...]


def canonical_pr_number(value: str) -> str:
    """Return a comparison key while leaving the display value untouched."""

    normalized = unicodedata.normalize("NFKC", value).strip().upper()
    normalized = re.sub(r"^PR[\s_./-]*", "", normalized)
    key = re.sub(r"[\s_./-]+", "", normalized)
    if not key or not any(character.isdigit() for character in key):
        raise ValueError(f"invalid PR number: {value!r}")
    if not all(character.isalnum() for character in key):
        raise ValueError(f"invalid PR number: {value!r}")
    return key


def extract_pr_numbers(text: str) -> tuple[str, ...]:
    """Extract exact PR identifier spellings from human-visible text."""

    found: list[tuple[int, str]] = []
    occupied: list[tuple[int, int]] = []
    normalized_text = unicodedata.normalize("NFC", text)
    for expression in (_LABELLED_PR_RE, _PREFIXED_PR_RE):
        for match in expression.finditer(normalized_text):
            span = match.span("value")
            if any(span[0] < end and start < span[1] for start, end in occupied):
                continue
            value = match.group("value").rstrip(".,;:)]}")
            try:
                canonical_pr_number(value)
            except ValueError:
                continue
            found.append((span[0], value))
            occupied.append(span)
    found.extend(_contextual_pr_values(normalized_text))
    found.sort(key=lambda item: item[0])
    values: list[str] = []
    seen: set[str] = set()
    for _, value in found:
        exact_key = unicodedata.normalize("NFC", value)
        if exact_key not in seen:
            seen.add(exact_key)
            values.append(value)
    return tuple(values)


def _contextual_pr_values(text: str) -> list[tuple[int, str]]:
    """Read bare PR values from labelled lists and PR table columns."""

    found: list[tuple[int, str]] = []
    line_offset = 0
    lines = text.splitlines(keepends=True)

    # Inline lists such as ``PR 번호: 00123, 00456`` and vertically stacked
    # lists headed by a line containing only ``PR 번호``.
    for index, line_with_end in enumerate(lines):
        line = line_with_end.rstrip("\r\n")
        label = _PR_LABEL_RE.search(line)
        if label is None:
            line_offset += len(line_with_end)
            continue
        tail = line[label.end() :]
        for value in _bare_pr_values(tail.lstrip(" `|:：#=-")):
            found.append((line_offset + max(label.end(), line.find(value, label.end())), value))
        stripped_label_line = re.sub(r"[`|:#：=\-\s]", "", line[label.start() :])
        if _PR_LABEL_RE.fullmatch(line[label.start() :].strip(" `|:：#=-")) or stripped_label_line.casefold() in {
            "pr번호",
            "prno.",
            "prnumber",
            "pr#",
        }:
            following_offset = line_offset + len(line_with_end)
            for following in lines[index + 1 :]:
                candidate_line = following.rstrip("\r\n").strip(" `|\t")
                if not candidate_line:
                    following_offset += len(following)
                    continue
                values = _bare_pr_values(candidate_line)
                if not values:
                    break
                for value in values:
                    found.append((following_offset + following.find(value), value))
                following_offset += len(following)
        line_offset += len(line_with_end)

    # Markdown tables: locate the labelled column and parse the same cell from
    # every subsequent row.  This covers a common bare-numeric PR layout.
    markdown_rows: list[tuple[int, list[str]]] = []
    offset = 0
    for line_with_end in lines:
        line = line_with_end.rstrip("\r\n")
        if "|" in line:
            cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
            markdown_rows.append((offset, cells))
        else:
            markdown_rows.append((offset, []))
        offset += len(line_with_end)
    for row_index, (row_offset, cells) in enumerate(markdown_rows):
        label_columns = [
            column for column, cell in enumerate(cells) if _PR_LABEL_RE.search(cell)
        ]
        if not label_columns:
            continue
        for following_offset, following_cells in markdown_rows[row_index + 1 :]:
            if not following_cells:
                break
            if all(re.fullmatch(r"[:\-\s]+", cell or "-") for cell in following_cells):
                continue
            for column in label_columns:
                if column >= len(following_cells):
                    continue
                for value in _bare_pr_values(following_cells[column]):
                    found.append((following_offset, value))

    # Merged native tables may be rendered as HTML rather than GFM.
    if "<table" in text.casefold() and _PR_LABEL_RE.search(text):
        try:
            root = lxml_html.fragment_fromstring(text, create_parent="div")
            for table in root.xpath(".//table"):
                label_columns: set[int] = set()
                for cells in _expanded_html_rows(table):
                    for column, cell in enumerate(cells):
                        if _PR_LABEL_RE.search(cell):
                            label_columns.add(column)
                            continue
                        if column in label_columns:
                            for value in _bare_pr_values(cell):
                                found.append((max(0, text.find(value)), value))
        except (TypeError, ValueError):
            pass
    return found


def _expanded_html_rows(table: Any) -> list[list[str]]:
    """Expand HTML row/column spans so labelled columns cannot drift."""

    rows: list[list[str]] = []
    pending: dict[int, tuple[int, str]] = {}
    for row in table.xpath(".//tr"):
        grid = {column: value for column, (_, value) in pending.items()}
        next_pending = {
            column: (remaining - 1, value)
            for column, (remaining, value) in pending.items()
            if remaining > 1
        }
        column = 0
        for cell in row.xpath("./th|./td"):
            while column in grid:
                column += 1
            value = " ".join(cell.text_content().split())
            try:
                column_span = max(1, int(cell.get("colspan", "1")))
                row_span = max(1, int(cell.get("rowspan", "1")))
            except (TypeError, ValueError):
                column_span = row_span = 1
            for offset in range(column_span):
                target = column + offset
                grid[target] = value
                if row_span > 1:
                    next_pending[target] = (row_span - 1, value)
            column += column_span
        pending = next_pending
        width = max(grid, default=-1) + 1
        rows.append([grid.get(index, "") for index in range(width)])
    return rows


def _bare_pr_values(value: str) -> tuple[str, ...]:
    values: list[str] = []
    for part in re.split(r"[,;，；]", value):
        candidate = part.strip(" `|\t:：#=()[]{}")
        if not _BARE_PR_VALUE_RE.fullmatch(candidate):
            continue
        if not any(character.isdigit() for character in candidate):
            continue
        try:
            canonical_pr_number(candidate)
        except ValueError:
            continue
        values.append(candidate)
    return tuple(values)


def extract_pr_ledger(
    records: Sequence[Mapping[str, Any]],
) -> tuple[dict[str, Any], ...]:
    """Extract PR occurrences with their immutable local provenance."""

    ledger: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for record in records:
        citation = str(record.get("citation", ""))
        visible = "\n".join(
            (
                str(record.get("slide_title", "")),
                str(record.get("content", "")),
            )
        )
        for value in extract_pr_numbers(visible):
            occurrence_key = (citation, unicodedata.normalize("NFC", value))
            if occurrence_key in seen:
                continue
            seen.add(occurrence_key)
            ledger.append(
                {
                    "value": value,
                    "canonical_key": canonical_pr_number(value),
                    "citation": citation,
                    "slide_number": int(record["slide_number"]),
                    "element_id": str(record["element_id"]),
                    "content_sha256": str(record.get("content_sha256", "")),
                }
            )
    return tuple(ledger)


def find_pr_number_mutations(
    text: str,
    accepted_pr_numbers: Sequence[str],
    *,
    evidence_text: str = "",
) -> tuple[str, ...]:
    """Find PR-like tokens that preserve a PR's digits but mutate its prefix."""

    accepted = {unicodedata.normalize("NFC", value) for value in accepted_pr_numbers}
    digit_signatures = {
        "".join(character for character in unicodedata.normalize("NFKC", value) if character.isdigit())
        for value in accepted_pr_numbers
    }
    digit_signatures.discard("")
    known_evidence_tokens = {
        unicodedata.normalize("NFC", match.group("value").rstrip(".,;:)]}"))
        for match in _PR_LIKE_TOKEN_RE.finditer(evidence_text)
    }
    mutations: list[str] = []
    for match in _PR_LIKE_TOKEN_RE.finditer(text):
        value = match.group("value").rstrip(".,;:)]}")
        exact = unicodedata.normalize("NFC", value)
        digits = "".join(
            character
            for character in unicodedata.normalize("NFKC", value)
            if character.isdigit()
        )
        prefix = re.match(r"(?i)[A-Z]+", value)
        if (
            exact not in accepted
            and exact not in known_evidence_tokens
            and digits in digit_signatures
            and prefix is not None
            and prefix.group(0).upper().startswith("P")
        ):
            mutations.append(value)
    return tuple(dict.fromkeys(mutations))


def build_source_semantic(
    corpus_dir: str | Path,
    *,
    identity: SourceIdentity,
    backend: ChatBackend | Callable[..., str],
    output_dir: str | Path,
    config: SemanticConfig | None = None,
) -> SourceSemanticExport:
    """Build one audited semantic dossier from one parsed PPTX corpus."""

    corpus = Path(corpus_dir).resolve()
    provenance_path = corpus / "provenance.jsonl"
    if not provenance_path.is_file():
        raise FileNotFoundError(provenance_path)
    destination = Path(output_dir).resolve()
    if destination.exists() and any(destination.iterdir()):
        raise ValueError(f"source semantic output directory is not empty: {destination}")
    destination.mkdir(parents=True, exist_ok=True)

    settings = config or SemanticConfig(coverage_policy="selected")
    records = load_provenance(provenance_path)
    if not records:
        raise ValueError("the parsed provenance corpus contains no evidence records")
    citations = tuple(str(record["citation"]) for record in records)
    if len(set(citations)) != len(citations):
        raise ValueError("parsed provenance contains duplicate citations")
    record_by_citation = {str(record["citation"]): record for record in records}

    ledger = extract_pr_ledger(records)
    if not ledger:
        raise ValueError(
            f"PR 번호를 찾지 못했습니다: {identity.source_name}; "
            "PR 번호가 이미지에만 있으면 OCR/VL 단계를 활성화하세요"
        )
    detected_prs = _ordered_pr_values(ledger)
    if identity.pr_numbers:
        expected_keys = tuple(canonical_pr_number(value) for value in identity.pr_numbers)
        detected_keys = tuple(canonical_pr_number(value) for value in detected_prs)
        if len(set(expected_keys)) != len(expected_keys):
            raise ValueError(
                f"source identity contains duplicate canonical PR numbers: "
                f"{identity.source_id}"
            )
        if expected_keys != detected_keys:
            raise ValueError(
                f"PR number inventory changed for {identity.source_id}: "
                f"expected {identity.pr_numbers!r}, detected {detected_prs!r}"
            )
        pr_numbers = identity.pr_numbers
    else:
        pr_numbers = detected_prs

    warnings: list[str] = []
    if settings.discover_topics:
        topics = _discover_topics(records, backend, settings, warnings)
    else:
        topics = _slide_topics(records)
    model_selected = {
        citation
        for topic in topics
        for citation in topic.citations
        if citation in record_by_citation
    }
    protected = {str(item["citation"]) for item in ledger}
    selected = model_selected | protected
    if settings.coverage_policy == "complete":
        selected.update(citations)
    if not selected:
        selected = set(citations)
        warnings.append("semantic selection was empty; all evidence was retained")

    duplicate_of = _exact_duplicate_map(records, protected)
    selected.update(
        original
        for duplicate, original in duplicate_of.items()
        if duplicate in selected
    )
    selected.difference_update(duplicate_of)
    selected.update(protected)
    selected_citations = tuple(citation for citation in citations if citation in selected)
    omitted_citations = tuple(citation for citation in citations if citation not in selected)
    selected_records = [record_by_citation[citation] for citation in selected_citations]
    if not selected_records:
        raise ValueError("source semantic selection contains no evidence")

    title = _source_title(identity.source_name, pr_numbers)
    generated, fallback, page_warnings = _generate_topic_page(
        Topic(title, selected_citations, settings.goal),
        selected_records,
        backend,
        settings,
    )
    warnings.extend(page_warnings)
    generated_body = _without_top_heading(generated).strip()
    invalid_prs = _unapproved_pr_spellings(
        generated_body,
        pr_numbers,
        evidence_text="\n".join(str(record.get("content", "")) for record in selected_records),
    )
    if invalid_prs:
        warnings.append(
            "model output changed or invented PR number(s); used verbatim fallback: "
            + ", ".join(invalid_prs)
        )
        generated_body = _without_top_heading(
            _verbatim_evidence_page(title, selected_records, set(selected_citations))
        ).strip()
        fallback = True
    if not fallback:
        try:
            _validate_source_line_grounding(
                generated_body,
                record_by_citation=record_by_citation,
                ledger=ledger,
            )
        except ValueError as error:
            warnings.append(
                f"model output borrowed a value/PR from another citation; used verbatim fallback: {error}"
            )
            generated_body = _without_top_heading(
                _verbatim_evidence_page(title, selected_records, set(selected_citations))
            ).strip()
            fallback = True

    body = _assemble_body(
        generated_body,
        pr_numbers=pr_numbers,
        ledger=ledger,
        selected_citations=selected_citations,
    )
    content_digest = sha256(body.encode("utf-8")).hexdigest()
    document = {
        "id": identity.source_id,
        "title": title,
        "description": settings.goal,
        "body_markdown": body,
        "citations": list(selected_citations),
        "generation": "verbatim_fallback" if fallback else "model",
        "content_sha256": content_digest,
        "source_id": identity.source_id,
        "pr_numbers": list(pr_numbers),
    }
    documents_text = json.dumps(
        document, ensure_ascii=False, sort_keys=True
    ) + "\n"
    documents_path = destination / "documents.jsonl"
    _write_text(documents_path, documents_text)

    decisions: list[dict[str, Any]] = []
    prs_by_citation: dict[str, list[str]] = {}
    for item in ledger:
        prs_by_citation.setdefault(str(item["citation"]), []).append(str(item["value"]))
    for citation in citations:
        if citation in protected:
            disposition = "keep"
            reason = "protected_pr"
            duplicate = None
        elif citation in duplicate_of:
            disposition = "duplicate"
            reason = "exact_duplicate"
            duplicate = duplicate_of[citation]
        elif citation in selected:
            disposition = "keep"
            reason = "model_selected"
            duplicate = None
        else:
            disposition = "omit"
            reason = "out_of_scope"
            duplicate = None
        decisions.append(
            {
                "citation": citation,
                "disposition": disposition,
                "reason_code": reason,
                "duplicate_of": duplicate,
                "pr_numbers": list(dict.fromkeys(prs_by_citation.get(citation, ()))),
            }
        )
    decisions_text = "".join(
        json.dumps(value, ensure_ascii=False, sort_keys=True) + "\n"
        for value in decisions
    )
    decisions_path = destination / "decisions.jsonl"
    _write_text(decisions_path, decisions_text)

    markdown_text = _render_semantic_markdown(
        document,
        source_name=identity.source_name,
    )
    markdown_path = destination / "semantic.md"
    _write_text(markdown_path, markdown_text)

    manifest = {
        "schema_version": SOURCE_SEMANTIC_SCHEMA_VERSION,
        "source_identity": {
            "source_id": identity.source_id,
            "source_name": identity.source_name,
            "source_sha256": identity.source_sha256,
            "pr_numbers": list(pr_numbers),
        },
        "source_provenance_sha256": sha256(provenance_path.read_bytes()).hexdigest(),
        "documents_file": documents_path.name,
        "documents_sha256": sha256(documents_text.encode("utf-8")).hexdigest(),
        "document_count": 1,
        "decisions_file": decisions_path.name,
        "decisions_sha256": sha256(decisions_text.encode("utf-8")).hexdigest(),
        "decision_count": len(decisions),
        "markdown_file": markdown_path.name,
        "markdown_sha256": sha256(markdown_text.encode("utf-8")).hexdigest(),
        "record_count": len(records),
        "selected_citations": list(selected_citations),
        "omitted_citations": list(omitted_citations),
        "pr_ledger": list(ledger),
        "fallback": fallback,
        "backend": {
            "type": type(backend).__name__,
            "model": getattr(backend, "model", None),
        },
        "config": asdict(settings),
        "warnings": warnings,
    }
    manifest_path = destination / "manifest.json"
    _write_text(
        manifest_path,
        json.dumps(manifest, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
    )
    # Re-load the complete artifact before returning.  Integration relies on
    # this same strict reader and therefore cannot observe a weaker contract.
    load_source_semantic(destination, corpus_dir=corpus)
    return SourceSemanticExport(
        output_dir=destination,
        manifest_path=manifest_path,
        documents_path=documents_path,
        decisions_path=decisions_path,
        markdown_path=markdown_path,
        source_id=identity.source_id,
        pr_numbers=tuple(pr_numbers),
        selected_citations=selected_citations,
        omitted_citations=omitted_citations,
        fallback=fallback,
        warnings=tuple(warnings),
    )


def load_source_semantic(
    path: str | Path,
    *,
    corpus_dir: str | Path | None = None,
) -> dict[str, Any]:
    """Load and verify a collection source-semantic artifact."""

    root = Path(path).resolve()
    manifest_path = root / "manifest.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FileNotFoundError(manifest_path) from None
    if not isinstance(manifest, Mapping):
        raise ValueError("source semantic manifest must be an object")
    if manifest.get("schema_version") != SOURCE_SEMANTIC_SCHEMA_VERSION:
        raise ValueError("unsupported source semantic schema")

    identity = manifest.get("source_identity")
    if not isinstance(identity, Mapping):
        raise ValueError("source semantic identity is missing")
    source_id = identity.get("source_id")
    if not isinstance(source_id, str) or not _safe_id(source_id):
        raise ValueError("source semantic source_id is unsafe")
    source_digest = identity.get("source_sha256")
    if not isinstance(source_digest, str) or not _DIGEST_RE.fullmatch(source_digest):
        raise ValueError("source semantic source_sha256 is invalid")
    pr_numbers = identity.get("pr_numbers")
    if (
        not isinstance(pr_numbers, list)
        or not pr_numbers
        or not all(isinstance(value, str) and value for value in pr_numbers)
    ):
        raise ValueError("source semantic pr_numbers must be a non-empty string array")
    for value in pr_numbers:
        canonical_pr_number(value)

    documents_raw = _verified_file(root, manifest, "documents_file", "documents_sha256")
    document_lines = [line for line in documents_raw.decode("utf-8").splitlines() if line.strip()]
    if len(document_lines) != 1 or manifest.get("document_count") != 1:
        raise ValueError("source semantic must contain exactly one canonical document")
    document = json.loads(document_lines[0])
    if not isinstance(document, dict):
        raise ValueError("source semantic document must be an object")
    if document.get("id") != source_id or document.get("source_id") != source_id:
        raise ValueError("source semantic document identity mismatch")
    if document.get("pr_numbers") != pr_numbers:
        raise ValueError("source semantic document PR inventory mismatch")
    body = document.get("body_markdown")
    if not isinstance(body, str):
        raise ValueError("source semantic body_markdown must be a string")
    if sha256(body.encode("utf-8")).hexdigest() != document.get("content_sha256"):
        raise ValueError("source semantic document content SHA-256 mismatch")
    for value in pr_numbers:
        if value not in body:
            raise ValueError(f"source semantic document lost PR number: {value}")

    decisions_raw = _verified_file(root, manifest, "decisions_file", "decisions_sha256")
    decisions = [
        json.loads(line)
        for line in decisions_raw.decode("utf-8").splitlines()
        if line.strip()
    ]
    if len(decisions) != manifest.get("decision_count") or len(decisions) != manifest.get("record_count"):
        raise ValueError("source semantic decision count mismatch")
    decision_citations = [value.get("citation") for value in decisions if isinstance(value, Mapping)]
    if len(set(decision_citations)) != len(decision_citations):
        raise ValueError("source semantic decisions contain duplicate citations")
    allowed_dispositions = {"keep", "omit", "duplicate"}
    for decision in decisions:
        if not isinstance(decision, Mapping) or decision.get("disposition") not in allowed_dispositions:
            raise ValueError("source semantic decision is invalid")

    markdown_raw = _verified_file(root, manifest, "markdown_file", "markdown_sha256")
    markdown = markdown_raw.decode("utf-8")
    frontmatter, markdown_body = _split_frontmatter(markdown)
    if frontmatter.get("source_id") != source_id or frontmatter.get("pr_numbers") != pr_numbers:
        raise ValueError("source semantic Markdown frontmatter identity mismatch")
    expected_tail = f"# {document['title']}\n\n{body}".rstrip() + "\n"
    if markdown_body != expected_tail:
        raise ValueError("source semantic Markdown is not an exact view of documents.jsonl")

    ledger = manifest.get("pr_ledger")
    if not isinstance(ledger, list) or not ledger:
        raise ValueError("source semantic PR ledger is missing")
    ledger_values = _ordered_pr_values(
        [item for item in ledger if isinstance(item, Mapping)]
    )
    if tuple(pr_numbers) != ledger_values:
        raise ValueError("source semantic PR ledger inventory mismatch")

    corpus = (
        Path(corpus_dir).resolve()
        if corpus_dir is not None
        else (root.parent / "parsed" / "corpus").resolve()
    )
    if corpus.is_dir():
        provenance_path = corpus / "provenance.jsonl"
        provenance_raw = provenance_path.read_bytes()
        if sha256(provenance_raw).hexdigest() != manifest.get("source_provenance_sha256"):
            raise ValueError("source semantic provenance SHA-256 mismatch")
        records = load_provenance(provenance_path)
        if len(records) != manifest.get("record_count"):
            raise ValueError("source semantic provenance record count mismatch")
        record_citations = [str(record["citation"]) for record in records]
        for record in records:
            if sha256(str(record["content"]).encode("utf-8")).hexdigest() != record.get(
                "content_sha256"
            ):
                raise ValueError(
                    f"parsed provenance content digest mismatch: {record['citation']}"
                )
        selected = manifest.get("selected_citations")
        omitted = manifest.get("omitted_citations")
        if (
            not isinstance(selected, list)
            or not isinstance(omitted, list)
            or not all(isinstance(value, str) for value in [*selected, *omitted])
        ):
            raise ValueError("source semantic citation inventories are invalid")
        if selected != [value for value in record_citations if value in set(selected)]:
            raise ValueError("source semantic selected citation order mismatch")
        if omitted != [value for value in record_citations if value in set(omitted)]:
            raise ValueError("source semantic omitted citation order mismatch")
        if set(selected) & set(omitted) or set(selected) | set(omitted) != set(record_citations):
            raise ValueError("source semantic selected/omitted coverage mismatch")
        if document.get("citations") != selected:
            raise ValueError("source semantic document citations do not match manifest")
        if decision_citations != record_citations:
            raise ValueError("source semantic decisions do not follow provenance order")
        kept = [
            str(value["citation"])
            for value in decisions
            if value.get("disposition") == "keep"
        ]
        if kept != selected:
            raise ValueError("source semantic keep decisions do not match selected citations")
        recomputed_ledger = list(extract_pr_ledger(records))
        if ledger != recomputed_ledger:
            raise ValueError("source semantic PR ledger does not match provenance")

        parsed_manifest_path = corpus.parent / "manifest.json"
        if parsed_manifest_path.is_file():
            parsed_manifest = json.loads(parsed_manifest_path.read_text(encoding="utf-8"))
            if not isinstance(parsed_manifest, Mapping):
                raise ValueError("parsed manifest must be an object")
            parsed_source = parsed_manifest.get("source")
            parsed_files = parsed_manifest.get("files")
            if (
                not isinstance(parsed_source, Mapping)
                or parsed_source.get("sha256") != source_digest
                or not isinstance(parsed_files, Mapping)
                or parsed_files.get("provenance_sha256")
                != manifest.get("source_provenance_sha256")
            ):
                raise ValueError("source semantic identity does not match parsed manifest")
    elif corpus_dir is not None:
        raise FileNotFoundError(corpus)
    return {
        "root": root,
        "manifest": dict(manifest),
        "manifest_sha256": sha256(manifest_path.read_bytes()).hexdigest(),
        "document": document,
        "decisions": decisions,
        "markdown": markdown,
        "markdown_sha256": sha256(markdown_raw).hexdigest(),
    }


def _ordered_pr_values(ledger: Sequence[Mapping[str, Any]]) -> tuple[str, ...]:
    values: list[str] = []
    seen_keys: set[str] = set()
    for item in ledger:
        value = str(item["value"])
        key = canonical_pr_number(value)
        if key not in seen_keys:
            seen_keys.add(key)
            values.append(value)
    return tuple(values)


def _exact_duplicate_map(
    records: Sequence[Mapping[str, Any]], protected: set[str]
) -> dict[str, str]:
    first_by_digest: dict[str, str] = {}
    duplicates: dict[str, str] = {}
    for record in records:
        citation = str(record["citation"])
        normalized = unicodedata.normalize("NFKC", str(record.get("content", "")))
        normalized = re.sub(r"\s+", " ", normalized).strip()
        payload = json.dumps(
            {
                "kind": str(record.get("kind", "")),
                "content_format": str(record.get("content_format", "")),
                "content": normalized,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
        digest = sha256(payload.encode("utf-8")).hexdigest()
        previous = first_by_digest.get(digest)
        if previous is None:
            first_by_digest[digest] = citation
        elif citation not in protected:
            duplicates[citation] = previous
    return duplicates


def _unapproved_pr_spellings(
    body: str,
    accepted: Sequence[str],
    *,
    evidence_text: str,
) -> tuple[str, ...]:
    accepted_exact = {unicodedata.normalize("NFC", value) for value in accepted}
    changed = [
        value
        for value in extract_pr_numbers(body)
        if unicodedata.normalize("NFC", value) not in accepted_exact
    ]
    changed.extend(
        find_pr_number_mutations(
            body,
            accepted,
            evidence_text=evidence_text,
        )
    )
    return tuple(dict.fromkeys(changed))


def _source_title(source_name: str, pr_numbers: Sequence[str]) -> str:
    stem = Path(source_name).stem.strip() or "PPTX source"
    return f"{stem} ({', '.join(pr_numbers)})"


def _validate_source_line_grounding(
    body: str,
    *,
    record_by_citation: Mapping[str, Mapping[str, Any]],
    ledger: Sequence[Mapping[str, Any]],
) -> None:
    pr_by_citation: dict[str, set[str]] = {}
    for item in ledger:
        pr_by_citation.setdefault(str(item["citation"]), set()).add(
            unicodedata.normalize("NFC", str(item["value"]))
        )
    for line_number, line in enumerate(body.splitlines(), start=1):
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        citations = [match.group(0) for match in CITATION_RE.finditer(stripped)]
        if not citations:
            continue
        evidence = "\n".join(
            f"{record_by_citation[citation].get('slide_title', '')}\n"
            f"{record_by_citation[citation].get('content', '')}"
            for citation in citations
        )
        borrowed_numbers = _numeric_tokens(stripped) - _numeric_tokens(evidence)
        if borrowed_numbers:
            raise ValueError(
                f"line {line_number} has unsupported number(s): "
                + ", ".join(sorted(borrowed_numbers))
            )
        supported_prs = {
            value for citation in citations for value in pr_by_citation.get(citation, ())
        }
        borrowed_prs = [
            value
            for value in extract_pr_numbers(stripped)
            if unicodedata.normalize("NFC", value) not in supported_prs
        ]
        if borrowed_prs:
            raise ValueError(
                f"line {line_number} has PR number from another citation: "
                + ", ".join(borrowed_prs)
            )


def _assemble_body(
    generated_body: str,
    *,
    pr_numbers: Sequence[str],
    ledger: Sequence[Mapping[str, Any]],
    selected_citations: Sequence[str],
) -> str:
    first_citation: dict[str, str] = {}
    for item in ledger:
        first_citation.setdefault(
            unicodedata.normalize("NFC", str(item["value"])),
            str(item["citation"]),
        )
    lines = ["## PR 번호", ""]
    exact_values = tuple(
        dict.fromkeys(unicodedata.normalize("NFC", str(item["value"])) for item in ledger)
    )
    for value in exact_values:
        lines.append(
            f"- `{value}` {first_citation[unicodedata.normalize('NFC', value)]}"
        )
    if generated_body:
        lines.extend(("", "## 의미 기반 정리", "", generated_body))
    lines.extend(("", "## 근거 블록", ""))
    lines.extend(f"- {citation}" for citation in selected_citations)
    return "\n".join(lines).strip()


def _render_semantic_markdown(document: Mapping[str, Any], *, source_name: str) -> str:
    frontmatter = {
        "title": str(document["title"]),
        "description": str(document.get("description", "")),
        "source_id": str(document["source_id"]),
        "source_name": source_name,
        "pr_numbers": list(document["pr_numbers"]),
        "tags": ["source", "reliability-analysis"],
        "draft": False,
    }
    yaml_text = yaml.safe_dump(
        frontmatter,
        allow_unicode=True,
        sort_keys=False,
        default_flow_style=False,
    ).rstrip()
    return (
        f"---\n{yaml_text}\n---\n\n# {document['title']}\n\n"
        f"{str(document['body_markdown']).rstrip()}\n"
    )


def _split_frontmatter(markdown: str) -> tuple[Mapping[str, Any], str]:
    if not markdown.startswith("---\n"):
        raise ValueError("source semantic Markdown is missing frontmatter")
    marker = markdown.find("\n---\n", 4)
    if marker < 0:
        raise ValueError("source semantic Markdown frontmatter is unterminated")
    value = yaml.safe_load(markdown[4:marker])
    if not isinstance(value, Mapping):
        raise ValueError("source semantic Markdown frontmatter must be a mapping")
    return value, markdown[marker + 5 :].lstrip("\n")


def _verified_file(
    root: Path,
    manifest: Mapping[str, Any],
    name_key: str,
    digest_key: str,
) -> bytes:
    name = manifest.get(name_key)
    if (
        not isinstance(name, str)
        or not name
        or Path(name).name != name
        or "/" in name
        or "\\" in name
    ):
        raise ValueError(f"source semantic {name_key} is unsafe")
    digest = manifest.get(digest_key)
    if not isinstance(digest, str) or not _DIGEST_RE.fullmatch(digest):
        raise ValueError(f"source semantic {digest_key} is invalid")
    raw = (root / name).read_bytes()
    actual = sha256(raw).hexdigest()
    if actual != digest:
        raise ValueError(
            f"source semantic {name} SHA-256 mismatch: expected {digest}, got {actual}"
        )
    return raw


def _safe_id(value: str) -> bool:
    return bool(re.fullmatch(r"[a-z0-9](?:[a-z0-9._-]{0,126}[a-z0-9])?", value))


def _write_text(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="\n")


__all__ = [
    "SOURCE_SEMANTIC_SCHEMA_VERSION",
    "SourceIdentity",
    "SourceSemanticExport",
    "build_source_semantic",
    "canonical_pr_number",
    "extract_pr_ledger",
    "extract_pr_numbers",
    "find_pr_number_mutations",
    "load_source_semantic",
]
