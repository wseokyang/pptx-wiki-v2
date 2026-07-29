from __future__ import annotations

import json
from pathlib import Path

import pytest

from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord, TableCell, TableData
from pptx_wiki.synthesis import (
    GroundingError,
    SynthesisConfig,
    synthesize_wiki,
    validate_grounded_markdown,
)
from pptx_wiki.wiki_output import export_slide_corpus, load_provenance, render_table


def _table(element_id: str, y: int, value: str) -> Element:
    return Element(
        id=element_id,
        slide_number=1,
        kind="table",
        bbox=BBox(100, y, 500, 200),
        z_index=y,
        table=TableData(
            rows=2,
            columns=2,
            cells=[
                TableCell(0, 0, "항목"),
                TableCell(0, 1, "값"),
                TableCell(1, 0, value),
                TableCell(1, 1, "100"),
            ],
        ),
    )


def test_export_is_deterministic_and_never_merges_close_tables(tmp_path: Path) -> None:
    # Input order is deliberately reversed and the visual gap is only 1 EMU.
    deck = DeckRecord(
        source_path="sample.pptx",
        slide_width=1_000,
        slide_height=1_000,
        slides=[
            SlideRecord(
                number=1,
                width=1_000,
                height=1_000,
                title="인접 표",
                elements=[_table("bottom", 301, "B"), _table("top", 100, "A")],
            )
        ],
    )
    first = export_slide_corpus(deck, tmp_path / "first")
    second = export_slide_corpus(deck, tmp_path / "second")

    assert first.digest == second.digest
    assert first.provenance_path.read_bytes() == second.provenance_path.read_bytes()
    markdown = first.slide_paths[0].read_text(encoding="utf-8")
    assert markdown.count("<!-- BEGIN TABLE") == 2
    assert markdown.index("[slide-1#top]") < markdown.index("[slide-1#bottom]")
    records = load_provenance(first.provenance_path)
    assert [record["citation"] for record in records] == [
        "[slide-1#top]",
        "[slide-1#bottom]",
    ]
    assert records[0]["table"] != records[1]["table"]


def test_merged_cells_use_html_spans() -> None:
    table = TableData(
        rows=2,
        columns=2,
        cells=[
            TableCell(0, 0, "병합", column_span=2),
            TableCell(0, 1, "", is_merge_origin=False),
            TableCell(1, 0, "A"),
            TableCell(1, 1, "B"),
        ],
    )
    rendered, content_format = render_table(table)
    assert content_format == "html"
    assert 'colspan="2"' in rendered
    assert rendered.count("<td") == 3


def test_numeric_grounding_rejects_llm_arithmetic() -> None:
    with pytest.raises(GroundingError, match="300"):
        validate_grounded_markdown(
            "# 매출\n\n- 합계는 300입니다. [slide-1#table-a]",
            {"[slide-1#table-a]"},
            numeric_evidence="값은 100과 200입니다.",
        )


class _UngroundedBackend:
    def complete(self, messages, *, max_tokens: int, temperature: float) -> str:
        prompt = messages[-1]["content"]
        if "Return JSON only" in prompt:
            return json.dumps(
                {
                    "topics": [
                        {"title": "매출", "citations": ["[slide-1#text-1]"]}
                    ]
                },
                ensure_ascii=False,
            )
        return "# 매출\n\n- 매출은 999입니다. [slide-1#text-1]"


def test_synthesis_falls_back_to_verbatim_on_ungrounded_number(tmp_path: Path) -> None:
    deck = DeckRecord(
        source_path="sample.pptx",
        slide_width=1_000,
        slide_height=1_000,
        slides=[
            SlideRecord(
                number=1,
                width=1_000,
                height=1_000,
                title="매출",
                elements=[
                    Element(
                        id="text-1",
                        slide_number=1,
                        kind="text",
                        bbox=BBox(0, 0, 500, 100),
                        z_index=0,
                        text="매출은 100입니다.",
                    )
                ],
            )
        ],
    )
    corpus = export_slide_corpus(deck, tmp_path / "corpus")
    result = synthesize_wiki(
        corpus.output_dir,
        backend=_UngroundedBackend(),
        config=SynthesisConfig(repair_attempts=0),
    )

    assert result.fallback_pages
    page = result.topic_paths[0].read_text(encoding="utf-8")
    assert "매출은 100입니다." in page
    assert "999" not in page
    assert "[slide-1#text-1]" in page
