from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree


@dataclass(frozen=True, slots=True)
class ArchiveLimits:
    max_members: int = 20_000
    max_member_bytes: int = 512 * 1024 * 1024
    max_total_bytes: int = 4 * 1024 * 1024 * 1024
    max_compression_ratio: float = 2_000.0


def validate_pptx_archive(
    path: str | Path,
    limits: ArchiveLimits = ArchiveLimits(),
    *,
    reject_external_resources: bool = False,
) -> None:
    """Reject malformed or suspiciously expansive OOXML archives before parsing."""

    source = Path(path)
    if source.suffix.lower() != ".pptx":
        raise ValueError("only macro-free .pptx input is accepted")
    with source.open("rb") as handle:
        validate_pptx_stream(
            handle,
            filename=source.name,
            limits=limits,
            reject_external_resources=reject_external_resources,
        )


def validate_pptx_stream(
    stream: BinaryIO,
    *,
    filename: str,
    limits: ArchiveLimits = ArchiveLimits(),
    reject_external_resources: bool = False,
) -> None:
    """Validate PPTX bytes from an already opened, caller-owned handle."""

    if Path(filename).suffix.casefold() != ".pptx":
        raise ValueError("only macro-free .pptx input is accepted")
    try:
        stream.seek(0)
    except (AttributeError, OSError) as exc:
        raise ValueError("PPTX validation requires a seekable binary stream") from exc
    try:
        with ZipFile(stream) as archive:
            members = archive.infolist()
            if len(members) > limits.max_members:
                raise ValueError(f"PPTX contains too many archive members: {len(members)}")
            total = 0
            for member in members:
                if member.filename.startswith(("/", "\\")) or ".." in Path(member.filename).parts:
                    raise ValueError(f"unsafe archive member path: {member.filename}")
                if member.file_size > limits.max_member_bytes:
                    raise ValueError(f"oversized archive member: {member.filename}")
                total += member.file_size
                if total > limits.max_total_bytes:
                    raise ValueError("PPTX expands beyond the configured size limit")
                if member.compress_size > 0 and member.file_size / member.compress_size > limits.max_compression_ratio:
                    raise ValueError(f"suspicious compression ratio: {member.filename}")
            if "[Content_Types].xml" not in archive.namelist() or "ppt/presentation.xml" not in archive.namelist():
                raise ValueError("archive is not a valid PPTX package")
            if reject_external_resources:
                external = _external_resource_relationships(archive)
                if external:
                    preview = ", ".join(f"{part} -> {target}" for part, target in external[:3])
                    raise ValueError(
                        "PPTX contains externally linked media/OLE resources that are blocked before rendering: "
                        + preview
                    )
    except BadZipFile as exc:
        raise ValueError("file is not a valid ZIP-based PPTX") from exc


def _external_resource_relationships(archive: ZipFile) -> list[tuple[str, str]]:
    blocked: list[tuple[str, str]] = []
    for name in archive.namelist():
        if not name.endswith(".rels"):
            continue
        try:
            root = ElementTree.fromstring(archive.read(name))
        except ElementTree.ParseError as exc:
            raise ValueError(f"invalid OOXML relationships part: {name}") from exc
        for relationship in root:
            if relationship.attrib.get("TargetMode", "").casefold() != "external":
                continue
            relationship_type = relationship.attrib.get("Type", "")
            if relationship_type.rstrip("/").endswith("/hyperlink"):
                continue
            blocked.append((name, relationship.attrib.get("Target", "<missing target>")))
    return blocked
