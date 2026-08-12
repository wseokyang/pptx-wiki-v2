from __future__ import annotations

from pathlib import Path

from pptx_wiki.extract import extract_pptx
from pptx_wiki.wiki_output import export_slide_corpus, load_provenance


def _output_bytes(root: Path) -> dict[str, bytes]:
    return {
        path.relative_to(root).as_posix(): path.read_bytes()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def test_extract_keeps_adjacent_native_tables_and_merged_cells(complex_pptx, tmp_path: Path) -> None:
    path, spec = complex_pptx
    deck = extract_pptx(
        path, assets_dir=tmp_path / "assets", include_images=True
    )
    slide = deck.slides[0]
    by_name = {element.name: element for element in slide.elements}

    table_a = by_name[spec.table_a_name]
    table_b = by_name[spec.table_b_name]
    assert table_a.id != table_b.id
    assert table_a.kind == table_b.kind == "table"
    assert table_b.bbox.y - table_a.bbox.y2 == spec.table_gap_emu
    assert table_a.metadata["native_boundary_authoritative"] is True
    assert table_b.metadata["native_boundary_authoritative"] is True
    assert table_a.metadata["ocr_policy"] == table_b.metadata["ocr_policy"] == "never"

    assert table_a.table is not None and table_b.table is not None
    cells_a = {(cell.row, cell.column): cell for cell in table_a.table.cells}
    cells_b = {(cell.row, cell.column): cell for cell in table_b.table.cells}
    assert cells_a[(0, 0)].column_span == 4
    assert cells_a[(0, 0)].text == "표 A — 국내 매출"
    assert cells_b[(0, 0)].row_span == 2
    # In the interchange model, False means a covered/spanned cell. Ordinary
    # extracted cells must therefore remain renderable table origins.
    assert cells_a[(1, 0)].is_merge_origin is True
    assert cells_a[(1, 0)].text == "제품"
    assert (0, 1) not in cells_a  # covered by A's horizontal merge
    assert (1, 0) not in cells_b  # covered by B's vertical merge


def test_extract_ignores_embedded_pictures_and_preserves_group_boundaries(
    complex_pptx, tmp_path: Path
) -> None:
    path, spec = complex_pptx
    assets_dir = tmp_path / "assets"
    deck = extract_pptx(path, assets_dir=assets_dir)
    by_name = {element.name: element for element in deck.slides[0].elements}

    assert spec.image_table_name not in by_name
    assert not any(
        element.kind in {"image", "picture"}
        for slide in deck.slides
        for element in slide.elements
    )
    assert not list((assets_dir / "images").glob("*"))

    group = by_name[spec.group_name]
    children = [element for element in deck.slides[0].elements if element.parent_id == group.id]
    assert {element.name for element in children} == {"GROUP_HEADING", "GROUP_VALUE"}
    assert all(element.bbox.x >= group.bbox.x for element in children)
    assert all(element.bbox.x2 <= group.bbox.x2 for element in children)


def test_fixture_extract_to_wiki_is_stable_and_never_joins_tables(complex_pptx, tmp_path: Path) -> None:
    path, _ = complex_pptx
    deck = extract_pptx(path, assets_dir=tmp_path / "assets")
    first = export_slide_corpus(deck, tmp_path / "wiki-a", document_id="close-tables-ko")
    second = export_slide_corpus(deck, tmp_path / "wiki-b", document_id="close-tables-ko")

    assert _output_bytes(first.output_dir) == _output_bytes(second.output_dir)
    slide_markdown = first.slide_paths[0].read_text(encoding="utf-8")
    assert slide_markdown.count("<!-- BEGIN TABLE") == 2
    assert slide_markdown.count("<!-- END TABLE") == 2
    assert "![" not in slide_markdown
    assert "<img" not in slide_markdown.casefold()
    assert "표 A — 국내 매출" in slide_markdown
    assert "제품" in slide_markdown
    assert "표 B" in slide_markdown
    assert "93.2%" in slide_markdown

    provenance = load_provenance(first.provenance_path)
    assert not any(
        record["kind"] in {"image", "picture"} or record["asset_path"]
        for record in provenance
    )
    table_records = [record for record in provenance if record["kind"] == "table"]
    assert [record["name"] for record in table_records] == [
        "NATIVE_TABLE_A",
        "NATIVE_TABLE_B",
    ]
    assert table_records[0]["element_id"] != table_records[1]["element_id"]
