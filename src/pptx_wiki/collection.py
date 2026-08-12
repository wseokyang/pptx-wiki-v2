"""One-shot multi-PPTX collection pipeline.

The existing single-file pipeline remains byte-compatible.  This orchestrator
adds a collection boundary around it and produces, in order:

``parsed evidence -> one semantic dossier per deck -> integrated artifact -> Quartz``.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from hashlib import sha256
import json
import os
from pathlib import Path
import stat
from typing import Any, Callable, Mapping, Sequence
import unicodedata

from .integration import (
    IntegratedExport,
    IntegrationConfig,
    build_integrated_artifact,
)
from .ocr import OCRAdapter
from .pipeline import PipelineConfig, PipelineResult, run_pipeline
from .safety import validate_pptx_archive, validate_pptx_stream
from .semantic import SemanticConfig
from .source_semantic import (
    SourceIdentity,
    SourceSemanticExport,
    build_source_semantic,
    canonical_pr_number,
    extract_pr_ledger,
)
from .synthesis import ChatBackend
from .wiki_output import load_provenance


COLLECTION_SCHEMA_VERSION = "pptx-wiki.collection.v1"
DEFAULT_COLLECTION_GOAL = (
    "반도체 패키지 신뢰성 의뢰·분석 결과에 직접 관련된 PR 번호, 제품·패키지, "
    "시료·Lot, 시험 방법·조건·표준·장비, 측정·관찰 결과, 불량 모드, 원인 판단, "
    "결론과 후속 조치만 보존하고 PPT 작성 가이드, 템플릿, 샘플 문구, 무관한 "
    "예시와 반복 boilerplate는 제거합니다. PR 번호는 원문 그대로 보존합니다."
)


@dataclass(frozen=True, slots=True)
class CollectionConfig:
    pipeline: PipelineConfig = field(default_factory=PipelineConfig)
    semantic: SemanticConfig = field(
        default_factory=lambda: SemanticConfig(
            goal=DEFAULT_COLLECTION_GOAL,
            coverage_policy="selected",
        )
    )
    integration: IntegrationConfig = field(default_factory=IntegrationConfig)
    recursive: bool = False
    site_title: str = "신뢰성 분석 LLM Wiki"
    max_files: int = 500
    max_total_bytes: int = 4 * 1024 * 1024 * 1024

    def __post_init__(self) -> None:
        if not self.site_title.strip() or "\n" in self.site_title or "\r" in self.site_title:
            raise ValueError("collection site_title must be one non-empty line")
        if self.max_files < 1:
            raise ValueError("collection max_files must be positive")
        if self.max_total_bytes < 1:
            raise ValueError("collection max_total_bytes must be positive")


@dataclass(frozen=True, slots=True)
class InputOccurrence:
    path: Path
    size: int
    mtime_ns: int
    sha256: str


@dataclass(frozen=True, slots=True)
class CollectionSourceResult:
    source_id: str
    source_dir: Path
    source_sha256: str
    occurrences: tuple[InputOccurrence, ...]
    pr_numbers: tuple[str, ...]
    parsed: PipelineResult
    semantic: SourceSemanticExport


@dataclass(frozen=True, slots=True)
class CollectionResult:
    output_dir: Path
    manifest_path: Path
    sources: tuple[CollectionSourceResult, ...]
    integrated: IntegratedExport
    quartz: Any
    input_count: int
    unique_source_count: int
    pr_numbers: tuple[str, ...]


@dataclass(slots=True)
class _PendingSource:
    source_id: str
    source_sha256: str
    occurrences: tuple[InputOccurrence, ...]
    source_dir: Path
    parsed: PipelineResult | None = None
    pr_numbers: tuple[str, ...] = ()
    semantic: SourceSemanticExport | None = None


def discover_pptx_inputs(
    inputs: Sequence[str | Path],
    *,
    recursive: bool = False,
    output_dir: str | Path | None = None,
    max_files: int = 500,
    max_total_bytes: int = 4 * 1024 * 1024 * 1024,
    reject_external_resources: bool = True,
) -> tuple[InputOccurrence, ...]:
    """Discover and fully preflight regular, non-reparse PPTX inputs."""

    if not inputs:
        raise ValueError("at least one PPTX file or directory is required")
    raw_roots = [Path(value).expanduser().absolute() for value in inputs]
    destination = Path(output_dir).expanduser().absolute() if output_dir is not None else None
    candidates: list[Path] = []
    seen_paths: set[str] = set()
    for raw in raw_roots:
        _reject_reparse_path(raw)
        if not raw.exists():
            raise FileNotFoundError(raw)
        if raw.is_dir():
            if destination is not None:
                _reject_containment(raw.resolve(), destination.resolve())
            values = _walk_pptx(raw, recursive=recursive)
        elif raw.is_file():
            values = [raw]
        else:
            raise ValueError(f"input is not a regular file or directory: {raw}")
        for candidate in values:
            _reject_reparse_path(candidate)
            if candidate.name.startswith("~$"):
                continue
            if candidate.suffix.casefold() != ".pptx":
                if raw.is_file():
                    raise ValueError(f"input file is not a .pptx: {candidate}")
                continue
            resolved = candidate.resolve()
            key = unicodedata.normalize("NFC", str(resolved)).casefold()
            if key not in seen_paths:
                seen_paths.add(key)
                candidates.append(resolved)
    candidates.sort(key=lambda path: unicodedata.normalize("NFC", str(path)).casefold())
    if not candidates:
        raise ValueError("no .pptx input files were found")
    if len(candidates) > max_files:
        raise ValueError(f"PPTX input count exceeds limit: {len(candidates)} > {max_files}")

    occurrences: list[InputOccurrence] = []
    total_bytes = 0
    for candidate in candidates:
        size, mtime_ns, digest = _inspect_input_handle(
            candidate,
            validate_archive=True,
            reject_external_resources=reject_external_resources,
        )
        total_bytes += size
        if total_bytes > max_total_bytes:
            raise ValueError(
                f"PPTX input size exceeds collection limit: {total_bytes} > {max_total_bytes}"
            )
        occurrences.append(
            InputOccurrence(
                path=candidate,
                size=size,
                mtime_ns=mtime_ns,
                sha256=digest,
            )
        )
    return tuple(occurrences)


def run_collection(
    inputs: Sequence[str | Path],
    output_dir: str | Path,
    *,
    semantic_backend: ChatBackend | Callable[..., str],
    integration_backend: ChatBackend | Callable[..., str] | None = None,
    ocr_adapter: OCRAdapter | None = None,
    config: CollectionConfig | None = None,
) -> CollectionResult:
    """Run all collection stages and return only after Quartz is published."""

    settings = config or CollectionConfig()
    lexical_destination = Path(output_dir).expanduser().absolute()
    _ensure_new_collection_output(lexical_destination)
    destination = lexical_destination.resolve()
    occurrences = discover_pptx_inputs(
        inputs,
        recursive=settings.recursive,
        output_dir=destination,
        max_files=settings.max_files,
        max_total_bytes=settings.max_total_bytes,
        reject_external_resources=settings.pipeline.block_external_resources,
    )
    grouped = _group_occurrences(occurrences)
    source_ids = _source_ids(tuple(grouped))
    pending = [
        _PendingSource(
            source_id=source_ids[digest],
            source_sha256=digest,
            occurrences=grouped[digest],
            source_dir=destination / "sources" / source_ids[digest],
        )
        for digest in sorted(grouped)
    ]

    _reject_reparse_path(lexical_destination)
    destination.mkdir(parents=True, exist_ok=True)
    if _is_reparse_point(destination):
        raise ValueError(f"collection output cannot be a reparse point: {lexical_destination}")
    lock_path = destination / ".collection.lock"
    with lock_path.open("x", encoding="utf-8") as handle:
        handle.write(str(os.getpid()))
    manifest_path = destination / "collection-manifest.json"
    _write_collection_manifest(
        manifest_path,
        status="running",
        occurrences=occurrences,
        pending=pending,
        config=settings,
    )
    integration = integration_backend or semantic_backend
    try:
        missing_pr_sources: list[str] = []
        # Parse every unique deck before the first LLM request.  This makes
        # stage 1 complete and lets PR preflight fail without uploading a deck.
        for item in pending:
            occurrence = item.occurrences[0]
            for alias in item.occurrences:
                _verify_unchanged(alias)
            item.source_dir.mkdir(parents=True)
            _write_text(item.source_dir / ".incomplete", "parsed\n")
            snapshot = _stage_verified_snapshot(
                occurrence,
                item.source_dir / ".input-snapshot",
                reject_external_resources=settings.pipeline.block_external_resources,
            )
            try:
                parsed = run_pipeline(
                    snapshot,
                    item.source_dir,
                    config=settings.pipeline,
                    ocr_adapter=ocr_adapter,
                    source_display_path=occurrence.path.name,
                )
                if _file_sha256(snapshot) != occurrence.sha256:
                    raise ValueError(
                        f"private input snapshot changed during parsing: {occurrence.path.name}"
                    )
            finally:
                if snapshot.is_file():
                    snapshot.unlink()
                snapshot_parent = snapshot.parent
                if snapshot_parent.is_dir() and not any(snapshot_parent.iterdir()):
                    snapshot_parent.rmdir()
            parsed_manifest = json.loads(
                parsed.parsed_manifest_path.read_text(encoding="utf-8")
            )
            parsed_source = parsed_manifest.get("source", {})
            if (
                not isinstance(parsed_source, Mapping)
                or parsed_source.get("sha256") != item.source_sha256
                or parsed_source.get("size") != occurrence.size
            ):
                raise ValueError(
                    f"parsed source identity does not match preflight: {occurrence.path}"
                )
            item.parsed = parsed
            records = load_provenance(parsed.corpus.provenance_path)
            ledger = extract_pr_ledger(records)
            pr_numbers = _ordered_pr_inventory(ledger)
            item.pr_numbers = pr_numbers
            if not pr_numbers:
                missing_pr_sources.append(occurrence.path.name)
            _write_source_identity(item)
        if missing_pr_sources:
            raise ValueError(
                "PR 번호를 찾지 못한 PPTX가 있습니다: "
                + ", ".join(missing_pr_sources)
                + ". 이미지 안의 PR은 기본적으로 무시됩니다; 네이티브 텍스트/표에 "
                "PR을 넣거나 extraction.include_images=true로 명시적으로 허용하세요"
            )

        for item in pending:
            assert item.parsed is not None
            _write_text(item.source_dir / ".incomplete", "semantic\n")
            item.semantic = build_source_semantic(
                item.parsed.corpus.output_dir,
                identity=SourceIdentity(
                    source_id=item.source_id,
                    source_name=item.occurrences[0].path.name,
                    source_sha256=item.source_sha256,
                    pr_numbers=item.pr_numbers,
                ),
                backend=semantic_backend,
                output_dir=item.source_dir / "semantic",
                config=settings.semantic,
            )
            _write_text(
                item.source_dir / ".complete.json",
                json.dumps(
                    {
                        "source_id": item.source_id,
                        "source_sha256": item.source_sha256,
                        "pr_numbers": list(item.pr_numbers),
                        "parsed_manifest_sha256": sha256(
                            item.parsed.parsed_manifest_path.read_bytes()
                        ).hexdigest(),
                        "semantic_manifest_sha256": sha256(
                            item.semantic.manifest_path.read_bytes()
                        ).hexdigest(),
                    },
                    ensure_ascii=False,
                    sort_keys=True,
                    indent=2,
                )
                + "\n",
            )
            (item.source_dir / ".incomplete").unlink()

        integrated = build_integrated_artifact(
            [item.semantic.output_dir for item in pending if item.semantic is not None],
            backend=integration,
            output_dir=destination / "integrated",
            config=settings.integration,
        )
        from .quartz_publish import publish_quartz

        quartz = publish_quartz(
            destination,
            integrated.output_dir,
            destination / "quartz",
            site_title=settings.site_title,
        )
        results = tuple(_source_result(item) for item in pending)
        all_prs = _all_pr_numbers(results)
        _write_deliverables_readme(destination, results, integrated)
        _write_collection_manifest(
            manifest_path,
            status="success",
            occurrences=occurrences,
            pending=pending,
            config=settings,
            integrated=integrated,
            quartz=quartz,
        )
        return CollectionResult(
            output_dir=destination,
            manifest_path=manifest_path,
            sources=results,
            integrated=integrated,
            quartz=quartz,
            input_count=len(occurrences),
            unique_source_count=len(results),
            pr_numbers=all_prs,
        )
    except Exception as error:
        _write_collection_manifest(
            manifest_path,
            status="failed",
            occurrences=occurrences,
            pending=pending,
            config=settings,
            error={
                "type": type(error).__name__,
                "message": _sanitise_error_message(
                    str(error), occurrences, destination
                ),
            },
        )
        raise
    finally:
        if lock_path.is_file():
            lock_path.unlink()


def _walk_pptx(directory: Path, *, recursive: bool) -> list[Path]:
    if not recursive:
        return [value for value in directory.iterdir() if value.is_file()]
    values: list[Path] = []
    for root, directory_names, file_names in os.walk(directory, followlinks=False):
        root_path = Path(root)
        safe_directories: list[str] = []
        for name in directory_names:
            child = root_path / name
            if _is_reparse_point(child):
                raise ValueError(f"input tree contains a reparse point: {child}")
            safe_directories.append(name)
        directory_names[:] = safe_directories
        values.extend(root_path / name for name in file_names)
    return values


def _reject_reparse_path(path: Path) -> None:
    current = path
    existing: list[Path] = []
    while True:
        if current.exists() or current.is_symlink():
            existing.append(current)
        if current.parent == current:
            break
        current = current.parent
    for value in reversed(existing):
        if _is_reparse_point(value):
            raise ValueError(f"symlink/junction/reparse paths are not accepted: {value}")


def _is_reparse_point(path: Path) -> bool:
    if path.is_symlink():
        return True
    is_junction = getattr(path, "is_junction", None)
    if callable(is_junction) and is_junction():
        return True
    try:
        attributes = getattr(os.lstat(path), "st_file_attributes", 0)
    except OSError:
        return False
    return bool(attributes & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))


def _reject_containment(input_dir: Path, output_dir: Path) -> None:
    if _is_relative_to(output_dir, input_dir) or _is_relative_to(input_dir, output_dir):
        raise ValueError(
            f"collection output and input directory must not contain each other: "
            f"input={input_dir}, output={output_dir}"
        )


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _group_occurrences(
    occurrences: Sequence[InputOccurrence],
) -> dict[str, tuple[InputOccurrence, ...]]:
    values: dict[str, list[InputOccurrence]] = {}
    for occurrence in occurrences:
        values.setdefault(occurrence.sha256, []).append(occurrence)
    return {key: tuple(value) for key, value in values.items()}


def _source_ids(digests: Sequence[str]) -> dict[str, str]:
    length = 16
    while True:
        prefixes = [value[:length] for value in digests]
        if len(set(prefixes)) == len(prefixes):
            return {value: f"deck-{value[:length]}" for value in digests}
        length += 2
        if length > 64:
            raise ValueError("unable to construct unique source ids")


def _ordered_pr_inventory(ledger: Sequence[Mapping[str, Any]]) -> tuple[str, ...]:
    values: list[str] = []
    seen: set[str] = set()
    for item in ledger:
        value = str(item["value"])
        key = canonical_pr_number(value)
        if key not in seen:
            seen.add(key)
            values.append(value)
    return tuple(values)


def _write_source_identity(item: _PendingSource) -> None:
    value = {
        "schema_version": "pptx-wiki.collection-source.v1",
        "source_id": item.source_id,
        "source_sha256": item.source_sha256,
        "source_name": item.occurrences[0].path.name,
        "pr_numbers": list(item.pr_numbers),
        "occurrences": [
            {
                "occurrence_id": _occurrence_id(occurrence.path),
                "name": occurrence.path.name,
                "size": occurrence.size,
                "mtime_ns": occurrence.mtime_ns,
                "sha256": occurrence.sha256,
            }
            for occurrence in item.occurrences
        ],
    }
    _write_text(
        item.source_dir / "source.json",
        json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
    )


def _verify_unchanged(occurrence: InputOccurrence) -> None:
    _reject_reparse_path(occurrence.path)
    size, mtime_ns, digest = _inspect_input_handle(
        occurrence.path,
        validate_archive=False,
        reject_external_resources=False,
    )
    if size != occurrence.size or mtime_ns != occurrence.mtime_ns:
        raise ValueError(f"input changed after preflight: {occurrence.path}")
    if digest != occurrence.sha256:
        raise ValueError(f"input content changed after preflight: {occurrence.path}")


def _stage_verified_snapshot(
    occurrence: InputOccurrence,
    snapshot_dir: Path,
    *,
    reject_external_resources: bool,
) -> Path:
    """Copy the exact preflight bytes once, then parse only that snapshot."""

    _reject_reparse_path(occurrence.path)
    snapshot_dir.mkdir(parents=True)
    snapshot = snapshot_dir / occurrence.path.name
    try:
        flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
        descriptor = os.open(occurrence.path, flags)
        try:
            opened = os.fstat(descriptor)
            identity_matches = _regular_path_identity_matches(
                opened,
                occurrence.path,
            )
            digest = sha256()
            with os.fdopen(descriptor, "rb", closefd=False) as source, snapshot.open("xb") as target:
                for chunk in iter(lambda: source.read(1024 * 1024), b""):
                    digest.update(chunk)
                    target.write(chunk)
                target.flush()
                os.fsync(target.fileno())
            after = os.fstat(descriptor)
            if (
                opened.st_size != after.st_size
                or opened.st_mtime_ns != after.st_mtime_ns
                or digest.hexdigest() != occurrence.sha256
            ):
                raise ValueError(f"input changed while staging: {occurrence.path.name}")
            if not identity_matches:
                _verify_current_path_digest(
                    occurrence.path,
                    occurrence.sha256,
                    error_message=(
                        f"input path changed while opening: {occurrence.path.name}"
                    ),
                )
        finally:
            os.close(descriptor)
        validate_pptx_archive(
            snapshot,
            reject_external_resources=reject_external_resources,
        )
        return snapshot
    except Exception:
        if snapshot.is_file():
            snapshot.unlink()
        if snapshot_dir.is_dir() and not any(snapshot_dir.iterdir()):
            snapshot_dir.rmdir()
        raise


def _inspect_input_handle(
    path: Path,
    *,
    validate_archive: bool,
    reject_external_resources: bool,
) -> tuple[int, int, str]:
    """Hash and optionally validate the same no-follow regular-file handle."""

    _reject_reparse_path(path)
    flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(path, flags)
    try:
        before = os.fstat(descriptor)
        identity_matches = _regular_path_identity_matches(before, path)
        with os.fdopen(descriptor, "rb", closefd=False) as handle:
            digest = sha256()
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
            if validate_archive:
                handle.seek(0)
                validate_pptx_stream(
                    handle,
                    filename=path.name,
                    reject_external_resources=reject_external_resources,
                )
        after = os.fstat(descriptor)
        if before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns:
            raise ValueError(f"input changed while reading: {path.name}")
        digest_value = digest.hexdigest()
        if not identity_matches:
            _verify_current_path_digest(
                path,
                digest_value,
                error_message=f"input path changed while opening: {path.name}",
            )
        return before.st_size, before.st_mtime_ns, digest_value
    finally:
        os.close(descriptor)


def _regular_path_identity_matches(opened: os.stat_result, path: Path) -> bool:
    """Compare handle/path identity, tolerating one transient provider result."""

    if not stat.S_ISREG(opened.st_mode):
        raise ValueError(f"input is not a regular file: {path.name}")
    for _ in range(2):
        current = os.lstat(path)
        if not stat.S_ISREG(current.st_mode):
            raise ValueError(f"input path is not a regular file: {path.name}")
        try:
            if os.path.samestat(opened, current):
                return True
        except (OSError, ValueError):
            pass
    return False


def _verify_current_path_digest(
    path: Path,
    expected_sha256: str,
    *,
    error_message: str,
) -> None:
    """Use an independent stable handle when a filesystem has unreliable IDs."""

    _reject_reparse_path(path)
    flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(path, flags)
    try:
        before = os.fstat(descriptor)
        if not stat.S_ISREG(before.st_mode):
            raise ValueError(error_message)
        digest = sha256()
        with os.fdopen(descriptor, "rb", closefd=False) as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        after = os.fstat(descriptor)
        if (
            before.st_size != after.st_size
            or before.st_mtime_ns != after.st_mtime_ns
            or digest.hexdigest() != expected_sha256
        ):
            raise ValueError(error_message)
    finally:
        os.close(descriptor)
    _reject_reparse_path(path)
    if not stat.S_ISREG(os.lstat(path).st_mode):
        raise ValueError(error_message)


def _source_result(item: _PendingSource) -> CollectionSourceResult:
    assert item.parsed is not None and item.semantic is not None
    return CollectionSourceResult(
        source_id=item.source_id,
        source_dir=item.source_dir,
        source_sha256=item.source_sha256,
        occurrences=item.occurrences,
        pr_numbers=item.pr_numbers,
        parsed=item.parsed,
        semantic=item.semantic,
    )


def _all_pr_numbers(sources: Sequence[CollectionSourceResult]) -> tuple[str, ...]:
    values: list[str] = []
    seen: set[str] = set()
    for source in sources:
        for value in source.pr_numbers:
            key = canonical_pr_number(value)
            if key not in seen:
                seen.add(key)
                values.append(value)
    return tuple(values)


def _write_collection_manifest(
    path: Path,
    *,
    status: str,
    occurrences: Sequence[InputOccurrence],
    pending: Sequence[_PendingSource],
    config: CollectionConfig,
    integrated: IntegratedExport | None = None,
    quartz: Any | None = None,
    error: Mapping[str, str] | None = None,
) -> None:
    value: dict[str, Any] = {
        "schema_version": COLLECTION_SCHEMA_VERSION,
        "status": status,
        "input_count": len(occurrences),
        "unique_source_count": len(pending),
        "inputs": [
            {
                "occurrence_id": _occurrence_id(occurrence.path),
                "name": occurrence.path.name,
                "size": occurrence.size,
                "mtime_ns": occurrence.mtime_ns,
                "sha256": occurrence.sha256,
            }
            for occurrence in occurrences
        ],
        "sources": [
            {
                "source_id": item.source_id,
                "source_sha256": item.source_sha256,
                "pr_numbers": list(item.pr_numbers),
                "occurrence_count": len(item.occurrences),
                "parsed_manifest": (
                    item.parsed.parsed_manifest_path.relative_to(path.parent).as_posix()
                    if item.parsed is not None
                    else None
                ),
                "semantic_manifest": (
                    item.semantic.manifest_path.relative_to(path.parent).as_posix()
                    if item.semantic is not None
                    else None
                ),
            }
            for item in pending
        ],
        "config": {
            "pipeline": {
                key: value
                for key, value in asdict(config.pipeline).items()
                if key not in {"office_binary", "pdf_binary", "scrub_env_vars"}
            },
            "semantic": asdict(config.semantic),
            "integration": asdict(config.integration),
            "recursive": config.recursive,
            "site_title": config.site_title,
            "max_files": config.max_files,
            "max_total_bytes": config.max_total_bytes,
        },
    }
    if integrated is not None:
        value["integrated"] = {
            "manifest": integrated.manifest_path.relative_to(path.parent).as_posix(),
            "manifest_sha256": sha256(integrated.manifest_path.read_bytes()).hexdigest(),
            "entity_count": integrated.entity_count,
            "relationship_count": integrated.relationship_count,
            "page_count": integrated.page_count,
        }
    if quartz is not None:
        quartz_manifest = Path(quartz.manifest_path)
        value["quartz"] = {
            "manifest": quartz_manifest.relative_to(path.parent).as_posix(),
            "manifest_sha256": sha256(quartz_manifest.read_bytes()).hexdigest(),
            "content": Path(quartz.content_dir).relative_to(path.parent).as_posix(),
            "page_count": quartz.page_count,
        }
        readme_path = path.parent / "README.md"
        if readme_path.is_file():
            value["deliverables"] = {
                "readme": readme_path.name,
                "readme_sha256": sha256(readme_path.read_bytes()).hexdigest(),
                "parsed": "sources/<source-id>/parsed/corpus/slides/*.md",
                "semantic": "sources/<source-id>/semantic/semantic.md",
                "knowledge_graph": (
                    "integrated/entities.jsonl + integrated/relationships.jsonl"
                    if integrated.relationships_path is not None
                    else "integrated/entities.jsonl"
                ),
                "quartz_content": "quartz/content",
            }
    if error is not None:
        value["error"] = dict(error)
    _write_text(
        path,
        json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
    )


def _ensure_new_collection_output(path: Path) -> None:
    _reject_reparse_path(path.parent)
    if path.exists() and _is_reparse_point(path):
        raise ValueError(f"collection output cannot be a reparse point: {path}")
    if path.exists() and not path.is_dir():
        raise ValueError(f"collection output exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()):
        raise ValueError(f"collection output directory is not empty: {path}")


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _write_deliverables_readme(
    destination: Path,
    sources: Sequence[CollectionSourceResult],
    integrated: IntegratedExport,
) -> None:
    lines = [
        "# PPTX collection deliverables",
        "",
        "## 1. 원본 파싱 Markdown",
        "",
    ]
    for source in sources:
        label = " / ".join(source.pr_numbers)
        first_slide = (
            f"sources/{source.source_id}/parsed/corpus/slides/slide-0001.md"
        )
        lines.append(f"- [{label} · parsed slides]({first_slide})")
    lines.extend(("", "## 2. 의미 기반 재구성 Markdown", ""))
    for source in sources:
        label = " / ".join(source.pr_numbers)
        semantic = f"sources/{source.source_id}/semantic/semantic.md"
        lines.append(f"- [{label} · semantic.md]({semantic})")
    lines.extend(("", "## 검증된 Knowledge Graph", ""))
    lines.append("- [KG nodes](integrated/entities.jsonl)")
    if integrated.relationships_path is not None:
        lines.append("- [KG relationships](integrated/relationships.jsonl)")
        lines.append(
            "- 모든 노드와 관계는 원본 evidence citation 및 PR lineage를 포함합니다."
        )
    else:
        lines.append("- `kg_profile: none`이므로 relationship edge는 생성하지 않았습니다.")
    lines.extend(
        (
            "",
            "## 3. Quartz용 output",
            "",
            "- [Quartz content index](quartz/content/index.md)",
            "- `quartz/content/`를 Quartz 프로젝트의 `content/`로 사용하세요.",
            "",
            "`integrated/`와 각 JSONL/manifest는 인용·해시·LLM 판단을 검증하기 위한 감사 산출물입니다.",
            "",
        )
    )
    _write_text(destination / "README.md", "\n".join(lines))


def _occurrence_id(path: Path) -> str:
    normalized = unicodedata.normalize("NFC", str(path)).casefold()
    return "occurrence-" + sha256(normalized.encode("utf-8")).hexdigest()[:16]


def _sanitise_error_message(
    message: str,
    occurrences: Sequence[InputOccurrence],
    destination: Path,
) -> str:
    value = message.replace(str(destination), "<collection-output>")
    for occurrence in occurrences:
        value = value.replace(str(occurrence.path), occurrence.path.name)
    return value


def _write_text(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="\n")


__all__ = [
    "COLLECTION_SCHEMA_VERSION",
    "DEFAULT_COLLECTION_GOAL",
    "CollectionConfig",
    "CollectionResult",
    "CollectionSourceResult",
    "InputOccurrence",
    "discover_pptx_inputs",
    "run_collection",
]
