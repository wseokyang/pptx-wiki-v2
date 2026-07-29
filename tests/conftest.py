from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures import ComplexPptxSpec, build_complex_pptx


@pytest.fixture()
def complex_pptx(tmp_path: Path) -> tuple[Path, ComplexPptxSpec]:
    path = tmp_path / "close-tables-ko.pptx"
    spec = build_complex_pptx(path)
    return path, spec
