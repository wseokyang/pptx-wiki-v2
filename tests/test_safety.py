from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from pptx_wiki.safety import validate_pptx_archive


def test_rejects_non_pptx_extension(tmp_path: Path) -> None:
    path = tmp_path / "deck.pptm"
    path.write_bytes(b"not used")
    with pytest.raises(ValueError, match="macro-free"):
        validate_pptx_archive(path)


def test_rejects_archive_without_presentation_parts(tmp_path: Path) -> None:
    path = tmp_path / "deck.pptx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("hello.txt", "hello")
    with pytest.raises(ValueError, match="valid PPTX"):
        validate_pptx_archive(path)


def test_rejects_external_media_before_powerpoint_rendering(tmp_path: Path) -> None:
    path = tmp_path / "linked.pptx"
    relationships = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="https://example.invalid/secret.png" TargetMode="External"/>
</Relationships>"""
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("ppt/presentation.xml", "<p:presentation xmlns:p='urn:p'/>")
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", relationships)

    with pytest.raises(ValueError, match="externally linked media"):
        validate_pptx_archive(path, reject_external_resources=True)
