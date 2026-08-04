from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import re
from typing import Mapping, Sequence

import pytest

from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord
from pptx_wiki.semantic import (
    SemanticConfig,
    build_semantic_output,
    load_semantic_documents,
)
from pptx_wiki.wiki_output import export_slide_corpus


REVENUE_CITATION = "[slide-1#revenue]"
GUIDANCE_CITATION = "[slide-1#authoring-guidance]"
UNKNOWN_CITATION = "[slide-9#fabricated]"
_CITATION_RE = re.compile(r"\[slide-\d+#[^\]\s#]+\]")
_NUMBER_RE = re.compile(r"(?<![0-9A-Za-z])\d+(?:\.\d+)?%?(?![0-9A-Za-z])")


class _ScriptedBackend:
    model = "scripted-semantic-model"

    def __init__(self, draft: str) -> None:
        self.draft = draft
        self.calls: list[Sequence[Mapping[str, str]]] = []

    def complete(
        self,
        messages: Sequence[Mapping[str, str]],
        *,
        max_tokens: int,
        temperature: float,
    ) -> str:
        self.calls.append(messages)
        if len(self.calls) == 1:
            return json.dumps(
                {
                    "topics": [
                        {
                            "title": "매출 요약",
                            "description": "목표에 맞는 매출 근거",
                            "citations": [REVENUE_CITATION],
                        }
                    ]
                },
                ensure_ascii=False,
            )
        return self.draft


class _CrossCitationNumberBackend:
    model = "scripted-cross-citation-model"

    def __init__(self) -> None:
        self.calls = 0

    def complete(
        self,
        messages: Sequence[Mapping[str, str]],
        *,
        max_tokens: int,
        temperature: float,
    ) -> str:
        self.calls += 1
        if self.calls == 1:
            return json.dumps(
                {
                    "topics": [
                        {
                            "title": "Revenue 40",
                            "citations": [REVENUE_CITATION],
                        },
                        {
                            "title": "Revenue 100",
                            "citations": [REVENUE_CITATION],
                        },
                    ]
                }
            )
        return f"# Revenue 100\n\n- Revenue is 100. {REVENUE_CITATION}"


@pytest.fixture()
def semantic_corpus(tmp_path: Path) -> Path:
    slide = SlideRecord(
        number=1,
        width=1_000,
        height=1_000,
        title="분기 요약",
        elements=[
            Element(
                id="revenue",
                slide_number=1,
                kind="text",
                bbox=BBox(10, 10, 400, 100),
                z_index=0,
                text="매출은 100입니다.",
            ),
            Element(
                id="authoring-guidance",
                slide_number=1,
                kind="text",
                bbox=BBox(10, 150, 400, 100),
                z_index=1,
                text="작성 지침의 비용 예시는 40입니다.",
            ),
        ],
    )
    deck = DeckRecord(
        source_path="semantic-source.pptx",
        slide_width=slide.width,
        slide_height=slide.height,
        slides=[slide],
    )
    return export_slide_corpus(
        deck,
        tmp_path / "corpus",
        document_id="semantic-stage-fixture",
    ).output_dir


def _tree_bytes(root: Path) -> dict[str, bytes]:
    return {
        path.relative_to(root).as_posix(): path.read_bytes()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def _claim_numbers(markdown: str) -> set[str]:
    without_citations = _CITATION_RE.sub("", markdown)
    return set(_NUMBER_RE.findall(without_citations))


def test_selected_coverage_records_omissions_hashes_and_preserves_corpus(
    semantic_corpus: Path,
    tmp_path: Path,
) -> None:
    before = _tree_bytes(semantic_corpus)
    semantic_dir = tmp_path / "semantic"
    backend = _ScriptedBackend(
        f"# 매출 요약\n\n- 매출은 100입니다. {REVENUE_CITATION}"
    )

    result = build_semantic_output(
        semantic_corpus,
        backend=backend,
        output_dir=semantic_dir,
        config=SemanticConfig(
            goal="매출 사실만 보존하고 작성 지침은 제외합니다.",
            coverage_policy="selected",
            repair_attempts=0,
        ),
    )

    manifest_path = semantic_dir / "manifest.json"
    documents_path = semantic_dir / "documents.jsonl"
    assert result.manifest_path == manifest_path
    assert result.documents_path == documents_path
    assert manifest_path.is_file() and documents_path.is_file()

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["schema_version"] == "pptx-wiki.semantic.v1"
    assert manifest["selected_citations"] == [REVENUE_CITATION]
    assert manifest["omitted_citations"] == [GUIDANCE_CITATION]
    assert result.selected_citations == (REVENUE_CITATION,)
    assert result.omitted_citations == (GUIDANCE_CITATION,)
    assert manifest["source_provenance_sha256"] == sha256(
        (semantic_corpus / "provenance.jsonl").read_bytes()
    ).hexdigest()
    assert manifest["documents_sha256"] == sha256(documents_path.read_bytes()).hexdigest()

    documents = load_semantic_documents(semantic_dir)
    assert len(documents) == manifest["document_count"] == 1
    document = documents[0]
    assert document["generation"] == "model"
    assert document["citations"] == [REVENUE_CITATION]
    assert set(_CITATION_RE.findall(document["body_markdown"])) <= set(document["citations"])
    assert _claim_numbers(document["body_markdown"]) <= {"100"}
    assert document["content_sha256"] == sha256(
        document["body_markdown"].encode("utf-8")
    ).hexdigest()
    assert _tree_bytes(semantic_corpus) == before


def test_topic_title_numbers_are_grounded_only_by_selected_citations(
    semantic_corpus: Path,
    tmp_path: Path,
) -> None:
    result = build_semantic_output(
        semantic_corpus,
        backend=_CrossCitationNumberBackend(),
        output_dir=tmp_path / "semantic",
        config=SemanticConfig(
            goal="Select revenue evidence only.",
            coverage_policy="selected",
            repair_attempts=0,
        ),
    )

    assert [document.title for document in result.documents] == ["Revenue 100"]
    assert result.selected_citations == (REVENUE_CITATION,)
    assert result.omitted_citations == (GUIDANCE_CITATION,)
    assert any(
        "rejected title with ungrounded number" in warning
        and "Revenue 40" in warning
        for warning in result.warnings
    )


def test_complete_coverage_retains_unselected_evidence_as_fallback_document(
    semantic_corpus: Path,
    tmp_path: Path,
) -> None:
    backend = _ScriptedBackend(
        f"# 매출 요약\n\n- 매출은 100입니다. {REVENUE_CITATION}"
    )
    semantic_dir = tmp_path / "semantic"

    result = build_semantic_output(
        semantic_corpus,
        backend=backend,
        output_dir=semantic_dir,
        config=SemanticConfig(
            goal="모든 원문 근거를 보존합니다.",
            coverage_policy="complete",
            repair_attempts=0,
        ),
    )

    manifest = json.loads((semantic_dir / "manifest.json").read_text(encoding="utf-8"))
    documents = load_semantic_documents(semantic_dir)
    covered = {
        citation
        for document in documents
        for citation in document["citations"]
    }
    assert covered == {REVENUE_CITATION, GUIDANCE_CITATION}
    assert manifest["omitted_citations"] == []

    fallback = next(
        document
        for document in documents
        if GUIDANCE_CITATION in document["citations"]
    )
    assert fallback["generation"] == "verbatim_fallback"
    assert "작성 지침의 비용 예시는 40입니다." in fallback["body_markdown"]
    assert GUIDANCE_CITATION in fallback["body_markdown"]
    assert fallback["id"] in manifest["fallback_documents"]
    assert fallback["id"] in result.fallback_documents


@pytest.mark.parametrize(
    ("draft", "forbidden"),
    [
        (
            f"# 매출 요약\n\n- 매출은 999입니다. {REVENUE_CITATION}",
            "999",
        ),
        (
            f"# 매출 요약\n\n- 매출은 100입니다. {UNKNOWN_CITATION}",
            UNKNOWN_CITATION,
        ),
    ],
)
def test_ungrounded_semantic_document_falls_back_to_verbatim_evidence(
    semantic_corpus: Path,
    tmp_path: Path,
    draft: str,
    forbidden: str,
) -> None:
    semantic_dir = tmp_path / "semantic"
    result = build_semantic_output(
        semantic_corpus,
        backend=_ScriptedBackend(draft),
        output_dir=semantic_dir,
        config=SemanticConfig(
            goal="매출 사실만 보존합니다.",
            coverage_policy="selected",
            repair_attempts=0,
        ),
    )

    document = load_semantic_documents(semantic_dir)[0]
    assert document["generation"] == "verbatim_fallback"
    assert document["citations"] == [REVENUE_CITATION]
    assert "매출은 100입니다." in document["body_markdown"]
    assert REVENUE_CITATION in document["body_markdown"]
    assert forbidden not in document["body_markdown"]
    assert result.fallback_documents == (document["id"],)
