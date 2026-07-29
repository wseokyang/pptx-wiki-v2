from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord
from pptx_wiki.validate import validate_deck


def test_numeric_conflict_between_native_and_ocr_is_an_error() -> None:
    native = Element("s001-e001", 1, "text", BBox(0, 0, 100, 100), 0, text="매출 1,200원")
    inferred = Element(
        "s001-e001-ocr",
        1,
        "ocr_text",
        BBox(0, 0, 100, 100),
        1,
        source="ocr",
        text="매출 1,700원",
    )
    slide = SlideRecord(1, 100, 100, elements=[native, inferred])
    deck = DeckRecord("fixture.pptx", 100, 100, [slide])

    issues = validate_deck(deck)

    assert any(issue.code == "numeric-conflict" and issue.severity == "error" for issue in issues)
