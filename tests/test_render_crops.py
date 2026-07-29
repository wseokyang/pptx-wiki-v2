from pathlib import Path

from PIL import Image

from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord
from pptx_wiki.render import create_element_crops


def test_create_crops_keeps_adjacent_pictures_separate(tmp_path: Path) -> None:
    slide_image = tmp_path / "slide.png"
    Image.new("RGB", (1000, 500), "white").save(slide_image)
    left = Element("s001-e001", 1, "picture", BBox(100, 100, 300, 300), 0)
    right = Element("s001-e002", 1, "picture", BBox(410, 100, 300, 300), 1)
    slide = SlideRecord(1, 1000, 500, elements=[left, right])
    deck = DeckRecord("fixture.pptx", 1000, 500, [slide])

    crops = create_element_crops(deck, [slide_image], tmp_path / "crops", padding_ratio=0.1)

    left_box = left.metadata["ocr_crop_bbox_emu"]
    right_box = right.metadata["ocr_crop_bbox_emu"]
    assert left_box["x"] + left_box["width"] <= right_box["x"]
    assert set(crops) == {left.id, right.id}
    with Image.open(crops[left.id]) as left_image:
        source_width = left.metadata["ocr_crop_bbox_px"]["right"] - left.metadata["ocr_crop_bbox_px"]["left"]
        assert left_image.width == source_width + 48
