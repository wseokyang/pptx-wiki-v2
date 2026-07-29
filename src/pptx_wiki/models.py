from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal


SourceKind = Literal["native", "ocr", "vlm", "derived"]


@dataclass(slots=True)
class BBox:
    """A slide-space bounding box in English Metric Units (EMU)."""

    x: int
    y: int
    width: int
    height: int

    @property
    def x2(self) -> int:
        return self.x + self.width

    @property
    def y2(self) -> int:
        return self.y + self.height

    @property
    def area(self) -> int:
        return max(0, self.width) * max(0, self.height)

    def normalized(self, slide_width: int, slide_height: int) -> list[float]:
        if slide_width <= 0 or slide_height <= 0:
            raise ValueError("slide dimensions must be positive")
        return [
            round(self.x / slide_width, 6),
            round(self.y / slide_height, 6),
            round(self.x2 / slide_width, 6),
            round(self.y2 / slide_height, 6),
        ]

    def intersects(self, other: "BBox") -> bool:
        return self.x < other.x2 and other.x < self.x2 and self.y < other.y2 and other.y < self.y2

    def intersection_area(self, other: "BBox") -> int:
        width = max(0, min(self.x2, other.x2) - max(self.x, other.x))
        height = max(0, min(self.y2, other.y2) - max(self.y, other.y))
        return width * height


@dataclass(slots=True)
class TableCell:
    row: int
    column: int
    text: str = ""
    row_span: int = 1
    column_span: int = 1
    is_merge_origin: bool = True
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class TableData:
    rows: int
    columns: int
    cells: list[TableCell] = field(default_factory=list)


@dataclass(slots=True)
class Element:
    id: str
    slide_number: int
    kind: str
    bbox: BBox
    z_index: int
    source: SourceKind = "native"
    name: str | None = None
    text: str | None = None
    markdown: str | None = None
    html: str | None = None
    table: TableData | None = None
    asset_path: str | None = None
    confidence: float | None = None
    parent_id: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class SlideRecord:
    number: int
    width: int
    height: int
    title: str | None = None
    notes: str | None = None
    elements: list[Element] = field(default_factory=list)
    rendered_path: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class DeckRecord:
    source_path: str
    slide_width: int
    slide_height: int
    slides: list[SlideRecord] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def source_name(self) -> str:
        return Path(self.source_path).name

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
