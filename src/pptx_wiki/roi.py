from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from .models import BBox


@dataclass(slots=True)
class CropBox:
    """Pixel-space crop rectangle using half-open coordinates."""

    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return max(0, self.right - self.left)

    @property
    def height(self) -> int:
        return max(0, self.bottom - self.top)


def non_overlapping_padded_bbox(
    target: BBox,
    blockers: Iterable[BBox],
    *,
    slide_width: int,
    slide_height: int,
    padding_ratio: float = 0.02,
) -> BBox:
    """Pad one object but stop at mid-gaps shared with neighbouring objects.

    PPT object boundaries are treated as authoritative.  This prevents two close
    tables or pictures from entering the same OCR crop even when a layout model
    would otherwise merge them.
    """

    pad_x = max(0, round(slide_width * padding_ratio))
    pad_y = max(0, round(slide_height * padding_ratio))
    left = max(0, target.x - pad_x)
    top = max(0, target.y - pad_y)
    right = min(slide_width, target.x2 + pad_x)
    bottom = min(slide_height, target.y2 + pad_y)

    for other in blockers:
        if other is target or other.area == 0:
            continue

        vertical_overlap = min(target.y2, other.y2) - max(target.y, other.y)
        horizontal_overlap = min(target.x2, other.x2) - max(target.x, other.x)

        if vertical_overlap > 0:
            if other.x2 <= target.x:
                left = max(left, (other.x2 + target.x + 1) // 2)
            elif target.x2 <= other.x:
                right = min(right, (target.x2 + other.x) // 2)

        if horizontal_overlap > 0:
            if other.y2 <= target.y:
                top = max(top, (other.y2 + target.y + 1) // 2)
            elif target.y2 <= other.y:
                bottom = min(bottom, (target.y2 + other.y) // 2)

    left = min(left, target.x)
    top = min(top, target.y)
    right = max(right, target.x2)
    bottom = max(bottom, target.y2)
    return BBox(left, top, right - left, bottom - top)


def emu_to_pixels(box: BBox, *, slide_width: int, slide_height: int, image_width: int, image_height: int) -> CropBox:
    if min(slide_width, slide_height, image_width, image_height) <= 0:
        raise ValueError("slide and image dimensions must be positive")
    left = max(0, min(image_width, box.x * image_width // slide_width))
    top = max(0, min(image_height, box.y * image_height // slide_height))
    right = max(left, min(image_width, (box.x2 * image_width + slide_width - 1) // slide_width))
    bottom = max(top, min(image_height, (box.y2 * image_height + slide_height - 1) // slide_height))
    return CropBox(left, top, right, bottom)
