from __future__ import annotations

from pathlib import Path

from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord, TableCell, TableData
from pptx_wiki.wiki_output import export_slide_corpus, load_provenance


def _two_table_deck(source_path: Path) -> DeckRecord:
    table_a = TableData(
        rows=2,
        columns=3,
        cells=[
            TableCell(0, 0, "표 A — 국내 매출", column_span=3),
            TableCell(0, 1, is_merge_origin=False),
            TableCell(0, 2, is_merge_origin=False),
            TableCell(1, 0, "제품"),
            TableCell(1, 1, "수량"),
            TableCell(1, 2, "매출"),
        ],
    )
    table_b = TableData(
        rows=2,
        columns=3,
        cells=[
            TableCell(0, 0, "표 B", row_span=2),
            TableCell(0, 1, "목표"),
            TableCell(0, 2, "실적"),
            TableCell(1, 0, is_merge_origin=False),
            TableCell(1, 1, "1,000"),
            TableCell(1, 2, "98.7%"),
        ],
    )
    slide = SlideRecord(
        number=1,
        width=12_192_000,
        height=6_858_000,
        title="인접 표 회귀 테스트",
        elements=[
            Element(
                id="shape-5-table",
                slide_number=1,
                kind="table",
                bbox=BBox(500_000, 1_250_000, 6_900_000, 1_400_000),
                z_index=4,
                name="NATIVE_TABLE_A",
                table=table_a,
            ),
            Element(
                id="shape-6-table",
                slide_number=1,
                kind="table",
                # Only 1pt (12,700 EMU) after table A.
                bbox=BBox(500_000, 2_662_700, 6_900_000, 1_400_000),
                z_index=5,
                name="NATIVE_TABLE_B",
                table=table_b,
            ),
        ],
    )
    return DeckRecord(
        source_path=str(source_path),
        slide_width=slide.width,
        slide_height=slide.height,
        slides=[slide],
    )


def _relative_files(root: Path) -> dict[str, bytes]:
    return {
        path.relative_to(root).as_posix(): path.read_bytes()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def test_markdown_and_json_are_byte_stable_and_tables_stay_separate(
    complex_pptx,
    tmp_path: Path,
) -> None:
    source_path, _ = complex_pptx
    deck = _two_table_deck(source_path)

    first = export_slide_corpus(deck, tmp_path / "first", document_id="fixture-ko")
    second = export_slide_corpus(deck, tmp_path / "second", document_id="fixture-ko")

    assert _relative_files(first.output_dir) == _relative_files(second.output_dir)

    markdown = first.slide_paths[0].read_text(encoding="utf-8")
    assert markdown.count("<!-- BEGIN TABLE") == 2
    assert markdown.count("<!-- END TABLE") == 2
    first_end = markdown.index("<!-- END TABLE", markdown.index("표 A"))
    second_begin = markdown.index("<!-- BEGIN TABLE", first_end)
    assert first_end < second_begin < markdown.index("표 B", second_begin)
    assert 'colspan="3"' in markdown
    assert 'rowspan="2"' in markdown

    records = load_provenance(first.provenance_path)
    assert [record["element_id"] for record in records] == [
        "shape-5-table",
        "shape-6-table",
    ]
    assert records[0]["content_sha256"] != records[1]["content_sha256"]
