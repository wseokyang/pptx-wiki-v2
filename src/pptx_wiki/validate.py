from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Literal

from .models import DeckRecord, Element


Severity = Literal["info", "warning", "error"]
_NUMBER = re.compile(r"(?<![0-9A-Za-z])[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:%|‰)?(?![0-9A-Za-z])")


@dataclass(slots=True)
class ValidationIssue:
    code: str
    message: str
    severity: Severity
    slide_number: int | None = None
    element_ids: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)


def _numbers(text: str | None) -> set[str]:
    return set(_NUMBER.findall(text or ""))


def validate_deck(deck: DeckRecord) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    seen: set[str] = set()
    for slide in deck.slides:
        by_id = {element.id: element for element in slide.elements}
        for element in slide.elements:
            if element.id in seen:
                issues.append(
                    ValidationIssue(
                        "duplicate-element-id",
                        f"duplicate element id: {element.id}",
                        "error",
                        slide.number,
                        [element.id],
                    )
                )

            if element.source in {"ocr", "vlm"} and "table" in element.kind.casefold() and element.parent_id:
                parent = by_id.get(element.parent_id)
                if parent is not None and parent.kind in {"image", "picture"}:
                    issues.append(
                        ValidationIssue(
                            "raster-table-boundary-unverified",
                            "table boundaries inside one bitmap are model-derived and require review when multiple tables may be present",
                            "warning",
                            slide.number,
                            [parent.id, element.id],
                        )
                    )
            seen.add(element.id)
            if element.bbox.x < 0 or element.bbox.y < 0 or element.bbox.x2 > slide.width or element.bbox.y2 > slide.height:
                issues.append(
                    ValidationIssue(
                        "element-outside-slide",
                        "element extends outside the slide canvas",
                        "warning",
                        slide.number,
                        [element.id],
                    )
                )
            if element.table:
                origins = [cell for cell in element.table.cells if cell.is_merge_origin]
                if any(cell.row >= element.table.rows or cell.column >= element.table.columns for cell in origins):
                    issues.append(
                        ValidationIssue(
                            "invalid-table-cell",
                            "table cell starts outside declared dimensions",
                            "error",
                            slide.number,
                            [element.id],
                        )
                    )
                if any(cell.row + cell.row_span > element.table.rows or cell.column + cell.column_span > element.table.columns for cell in origins):
                    issues.append(
                        ValidationIssue(
                            "invalid-table-span",
                            "rowspan or colspan exceeds table dimensions",
                            "error",
                            slide.number,
                            [element.id],
                        )
                    )

        native = [element for element in slide.elements if element.source == "native" and (element.text or element.table)]
        generated = [element for element in slide.elements if element.source in {"ocr", "vlm"}]
        for inferred in generated:
            for exact in native:
                overlap = inferred.bbox.intersection_area(exact.bbox)
                if overlap == 0 or min(inferred.bbox.area, exact.bbox.area) == 0:
                    continue
                coverage = overlap / min(inferred.bbox.area, exact.bbox.area)
                if coverage < 0.75:
                    continue
                exact_numbers = _numbers(exact.text)
                inferred_numbers = _numbers(inferred.text)
                if exact_numbers and inferred_numbers and exact_numbers != inferred_numbers:
                    issues.append(
                        ValidationIssue(
                            "numeric-conflict",
                            "OCR/VLM numbers disagree with overlapping native content",
                            "error",
                            slide.number,
                            [exact.id, inferred.id],
                            {"native": sorted(exact_numbers), "generated": sorted(inferred_numbers)},
                        )
                    )
    return issues


def issues_to_dicts(issues: list[ValidationIssue]) -> list[dict[str, Any]]:
    return [
        {
            "code": issue.code,
            "message": issue.message,
            "severity": issue.severity,
            "slide_number": issue.slide_number,
            "element_ids": issue.element_ids,
            "details": issue.details,
        }
        for issue in issues
    ]
