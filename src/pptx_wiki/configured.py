from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Sequence

from .collection import (
    DEFAULT_COLLECTION_GOAL,
    CollectionConfig,
    CollectionResult,
    run_collection,
)
from .config import AppConfig
from .integration import IntegrationConfig
from .ocr import (
    CommandOCRAdapter,
    FallbackOCRAdapter,
    OCRAdapter,
    OpenAICompatibleVLMAdapter,
    PaddleOCRCLIAdapter,
    PersistentOCRWorkerAdapter,
)
from .pipeline import PipelineConfig, PipelineResult, run_pipeline
from .semantic import SemanticConfig, build_semantic_output
from .synthesis import OpenAICompatibleClient, WikiSynthesis
from .wiki_publish import publish_wiki


def run_configured(
    pptx_path: str | Path,
    config: AppConfig,
    *,
    output_override: str | Path | None = None,
) -> PipelineResult:
    """Run one PPTX using a previously validated trusted configuration."""

    source = Path(pptx_path).expanduser().resolve()
    destination = (
        Path(output_override).expanduser().resolve()
        if output_override is not None
        else config.output.path_for(source)
    )
    _check_output(
        destination,
        config.output.allow_existing,
        stages=("parsed", "semantic", "wiki"),
    )
    adapter = build_ocr_adapter(config)
    semantic_backend = (
        OpenAICompatibleClient(
            base_url=config.llm_api.base_url,
            model=config.llm_api.model,
            api_key=config.llm_api.resolved_api_key(),
            timeout_seconds=config.llm_api.timeout_seconds,
        )
        if config.semantic.enabled
        else None
    )
    rendered_dir = config.render.rendered_slides_dir if adapter is not None else None
    try:
        result = run_pipeline(
            source,
            destination,
            config=PipelineConfig(
                render_backend=config.render.backend,
                dpi=config.render.dpi,
                source_padding_ratio=config.render.source_padding_ratio,
                model_padding_px=config.render.model_padding_px,
                include_empty_shapes=config.extraction.include_empty_shapes,
                strict_extraction=config.extraction.strict,
                strict_ocr=config.ocr.strict,
                office_binary=config.render.office_binary,
                pdf_binary=config.render.pdf_binary,
                scrub_env_vars=tuple(
                    dict.fromkeys(
                        name
                        for name in (config.vlm_api.api_key_env, config.llm_api.api_key_env)
                        if name
                    )
                ),
                block_external_resources=config.render.block_external_resources,
            ),
            ocr_adapter=adapter,
            rendered_slides_dir=rendered_dir,
        )
    finally:
        close = getattr(adapter, "close", None)
        if callable(close):
            close()

    semantic = None
    wiki = None
    if config.semantic.enabled:
        assert semantic_backend is not None
        semantic = build_semantic_output(
            result.corpus.output_dir,
            backend=semantic_backend,
            output_dir=destination / "semantic",
            config=SemanticConfig(
                goal=config.semantic.goal or SemanticConfig().goal,
                coverage_policy=config.semantic.coverage_policy,
                language=config.semantic.language,
                max_input_chars=config.semantic.max_input_chars,
                max_output_tokens=min(
                    config.semantic.max_output_tokens,
                    config.llm_api.max_tokens,
                ),
                max_topics=config.semantic.max_topics,
                repair_attempts=config.semantic.repair_attempts,
                discover_topics=config.semantic.discover_topics,
            ),
        )
    if config.wiki.enabled:
        if semantic is None:
            raise ValueError("wiki stage requires a semantic artifact")
        published = publish_wiki(
            semantic.output_dir,
            result.corpus.output_dir,
            destination / "wiki",
        )
        wiki = WikiSynthesis(
            output_dir=published.output_dir,
            index_path=published.index_path,
            report_path=published.report_path,
            topic_paths=published.page_paths,
            topic_count=published.page_count,
            fallback_pages=tuple(
                f"{document_id}.md" for document_id in semantic.fallback_documents
            ),
            warnings=semantic.warnings,
            semantic=semantic,
        )
    return replace(result, semantic=semantic, wiki=wiki)


def run_configured_collection(
    inputs: Sequence[str | Path],
    config: AppConfig,
    *,
    output_dir: str | Path,
    recursive: bool = False,
    site_title: str = "신뢰성 분석 LLM Wiki",
    max_entities: int = 256,
    max_files: int = 500,
    max_total_bytes: int = 4 * 1024 * 1024 * 1024,
) -> CollectionResult:
    """Run the trusted-config one-shot pipeline for multiple PPTX files."""

    if not config.semantic.enabled:
        raise ValueError("collection requires semantic.enabled=true")
    if not config.wiki.enabled:
        raise ValueError("collection requires wiki.enabled=true")
    if config.render.rendered_slides_dir is not None:
        raise ValueError(
            "render.rendered_slides_dir is ambiguous for a collection; "
            "disable it and let each PPTX be rendered independently"
        )
    destination = Path(output_dir).expanduser().absolute()
    adapter = build_ocr_adapter(config)
    backend = OpenAICompatibleClient(
        base_url=config.llm_api.base_url,
        model=config.llm_api.model,
        api_key=config.llm_api.resolved_api_key(),
        timeout_seconds=config.llm_api.timeout_seconds,
    )
    goal = config.semantic.goal or DEFAULT_COLLECTION_GOAL
    try:
        return run_collection(
            inputs,
            destination,
            semantic_backend=backend,
            integration_backend=backend,
            ocr_adapter=adapter,
            config=CollectionConfig(
                pipeline=PipelineConfig(
                    render_backend=config.render.backend,
                    dpi=config.render.dpi,
                    source_padding_ratio=config.render.source_padding_ratio,
                    model_padding_px=config.render.model_padding_px,
                    include_empty_shapes=config.extraction.include_empty_shapes,
                    strict_extraction=config.extraction.strict,
                    strict_ocr=config.ocr.strict,
                    office_binary=config.render.office_binary,
                    pdf_binary=config.render.pdf_binary,
                    scrub_env_vars=tuple(
                        dict.fromkeys(
                            name
                            for name in (
                                config.vlm_api.api_key_env,
                                config.llm_api.api_key_env,
                            )
                            if name
                        )
                    ),
                    block_external_resources=config.render.block_external_resources,
                ),
                semantic=SemanticConfig(
                    goal=goal,
                    coverage_policy=config.semantic.coverage_policy,
                    language=config.semantic.language,
                    max_input_chars=config.semantic.max_input_chars,
                    max_output_tokens=min(
                        config.semantic.max_output_tokens,
                        config.llm_api.max_tokens,
                    ),
                    max_topics=config.semantic.max_topics,
                    repair_attempts=config.semantic.repair_attempts,
                    discover_topics=config.semantic.discover_topics,
                ),
                integration=IntegrationConfig(
                    goal=goal,
                    language=config.semantic.language,
                    max_input_chars=config.semantic.max_input_chars,
                    max_output_tokens=min(
                        config.semantic.max_output_tokens,
                        config.llm_api.max_tokens,
                    ),
                    max_entities=max_entities,
                    max_topics=config.semantic.max_topics,
                    repair_attempts=config.semantic.repair_attempts,
                ),
                recursive=recursive,
                site_title=site_title,
                max_files=max_files,
                max_total_bytes=max_total_bytes,
            ),
        )
    finally:
        close = getattr(adapter, "close", None)
        if callable(close):
            close()


def build_ocr_adapter(config: AppConfig) -> OCRAdapter | None:
    settings = config.ocr
    if not settings.enabled or settings.backend == "none":
        return None
    scrub_env_vars = tuple(
        dict.fromkeys(
            name
            for name in (config.vlm_api.api_key_env, config.llm_api.api_key_env)
            if name
        )
    )

    def paddle() -> PaddleOCRCLIAdapter:
        value = settings.paddle
        return PaddleOCRCLIAdapter(
            executable=value.executable,
            pipeline_version=value.pipeline_version,
            device=value.device,
            engine=value.engine,
            timeout=value.timeout_seconds,
            extra_args=value.extra_args,
            scrub_env_vars=scrub_env_vars,
        )

    def command() -> CommandOCRAdapter:
        value = settings.command
        return CommandOCRAdapter(
            value.argv,
            cwd=value.cwd,
            timeout=value.timeout_seconds,
            scrub_env_vars=scrub_env_vars,
        )

    def vlm() -> OpenAICompatibleVLMAdapter:
        value = config.vlm_api
        return OpenAICompatibleVLMAdapter(
            base_url=value.base_url,
            model=value.model,
            api_key=value.resolved_api_key(),
            timeout=value.timeout_seconds,
            max_tokens=value.max_tokens,
            response_format=settings.response_format,
            image_detail=settings.image_detail,
        )

    def local_model() -> PersistentOCRWorkerAdapter:
        value = settings.local_model
        python_executable = value.resolved_python_executable()
        worker_script = value.resolved_worker_script
        model_directory = value.resolved_model_directory
        if not python_executable.is_file():
            raise ValueError(
                f"{value.model} environment is not installed: {python_executable}; "
                f"run workers/{value.model}/setup-windows.ps1"
            )
        if not worker_script.is_file():
            raise ValueError(f"bundled OCR worker is missing: {worker_script}")
        if not model_directory.is_dir():
            raise ValueError(
                f"{value.model} model is not downloaded: {model_directory}; "
                f"run workers/{value.model}/download.py in its virtual environment"
            )
        command_line = [
            str(python_executable),
            "-u",
            str(worker_script),
            "--serve",
            "--model-dir",
            str(model_directory),
            "--device",
            value.device,
            "--dtype",
            value.dtype,
            "--max-new-tokens",
            str(value.max_new_tokens),
        ]
        return PersistentOCRWorkerAdapter(
            command_line,
            backend_name=value.model,
            startup_timeout=value.startup_timeout_seconds,
            request_timeout=value.request_timeout_seconds,
            cwd=value.profile_directory,
            scrub_env_vars=scrub_env_vars,
        )

    if settings.backend == "openai_vlm":
        return vlm()
    if settings.backend == "local_model":
        adapter = local_model()
        return (
            FallbackOCRAdapter([adapter, vlm()])
            if settings.local_model.fallback_to_vlm
            else adapter
        )
    if settings.backend == "paddle_cli":
        return paddle()
    if settings.backend == "command":
        return command()
    if settings.backend == "local_then_vlm":
        return FallbackOCRAdapter([local_model(), vlm()])
    if settings.backend == "paddle_then_vlm":
        return FallbackOCRAdapter([paddle(), vlm()])
    if settings.backend == "command_then_vlm":
        return FallbackOCRAdapter([command(), vlm()])
    raise ValueError(f"unsupported OCR backend: {settings.backend}")


def _check_output(
    path: Path,
    allow_existing: bool,
    *,
    stages: tuple[str, ...] = (),
) -> None:
    if path.exists() and not path.is_dir():
        raise ValueError(f"output path exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()) and not allow_existing:
        raise ValueError(
            f"output directory is not empty: {path}; change output.directory or explicitly set output.allow_existing=true"
        )
    for stage in stages:
        target = path / stage
        if target.exists() and not target.is_dir():
            raise ValueError(f"{stage} output exists and is not a directory: {target}")
        if target.exists() and any(target.iterdir()):
            raise ValueError(f"{stage} output directory is not empty: {target}")


__all__ = ["build_ocr_adapter", "run_configured", "run_configured_collection"]
