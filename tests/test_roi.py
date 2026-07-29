from pptx_wiki.models import BBox
from pptx_wiki.roi import emu_to_pixels, non_overlapping_padded_bbox


def test_adjacent_table_crops_do_not_overlap() -> None:
    left_table = BBox(100, 100, 300, 500)
    right_table = BBox(410, 100, 300, 500)

    left_crop = non_overlapping_padded_bbox(
        left_table,
        [right_table],
        slide_width=1000,
        slide_height=800,
        padding_ratio=0.1,
    )
    right_crop = non_overlapping_padded_bbox(
        right_table,
        [left_table],
        slide_width=1000,
        slide_height=800,
        padding_ratio=0.1,
    )

    assert left_crop.x2 <= right_crop.x
    assert left_crop.x <= left_table.x
    assert left_crop.x2 >= left_table.x2
    assert right_crop.x <= right_table.x
    assert right_crop.x2 >= right_table.x2


def test_emu_to_pixels_clamps_to_image() -> None:
    box = BBox(-10, 50, 1_020, 760)
    crop = emu_to_pixels(box, slide_width=1000, slide_height=800, image_width=4000, image_height=3200)
    assert (crop.left, crop.top, crop.right, crop.bottom) == (0, 200, 4000, 3200)
