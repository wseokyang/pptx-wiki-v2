from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
from typing import Sequence

from .config import load_config
from .collection import DEFAULT_COLLECTION_GOAL
from .configured import run_configured, run_configured_collection
from .ocr import CommandOCRAdapter, FallbackOCRAdapter, OpenAICompatibleVLMAdapter, PaddleOCRCLIAdapter
from .pipeline import PipelineConfig, run_pipeline
from .quartz_publish import publish_quartz
from .semantic import SemanticConfig, build_semantic_output
from .synthesis import OpenAICompatibleClient
from .wiki_publish import publish_wiki


DEFAULT_SEMANTIC_GOAL = (
    "Retain substantive source content and exclude authoring guidance, "
    "templates, boilerplate, and unrelated examples."
)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-wiki",
        description="Parse PPTX evidence, reorganize it semantically, then publish a grounded Wiki.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    convert = subparsers.add_parser(
        "convert",
        help="run configured parsed, semantic, and Wiki stages",
    )
    convert.add_argument("input", type=Path)
    convert.add_argument("--config", required=True, type=Path)
    convert.add_argument("-o", "--output", type=Path)

    collection = subparsers.add_parser(
        "batch",
        aliases=["collection"],
        help="process multiple PPTX files through parsed, semantic, integrated, and Quartz stages",
    )
    collection.add_argument(
        "input",
        nargs="+",
        type=Path,
        help="one or more .pptx files or directories",
    )
    collection.add_argument("--config", required=True, type=Path)
    collection.add_argument("-o", "--output", required=True, type=Path)
    collection.add_argument("--recursive", action="store_true")
    collection.add_argument("--site-title", default="신뢰성 분석 LLM Wiki")
    collection.add_argument("--max-entities", type=int, default=256)
    collection.add_argument("--max-files", type=int, default=500)
    collection.add_argument("--max-total-mib", type=int, default=4096)

    parse = subparsers.add_parser("parse", help="create the source-faithful parsed artifact")
    _add_parse_arguments(parse, include_semantic=False)

    run = subparsers.add_parser(
        "run",
        help="compatibility shortcut for parse, optionally continuing through semantic and Wiki",
    )
    _add_parse_arguments(run, include_semantic=True)

    organize = subparsers.add_parser(
        "organize",
        aliases=["semantic"],
        help="create a grounded semantic artifact from parsed provenance",
    )
    organize.add_argument("input", type=Path, help="parsed directory or corpus directory")
    organize.add_argument("-o", "--output", required=True, type=Path)
    _add_semantic_arguments(organize, coverage_default="selected")

    wiki = subparsers.add_parser(
        "wiki",
        help="publish Wiki files deterministically from a semantic artifact",
    )
    wiki.add_argument("input", type=Path, help="semantic artifact directory")
    wiki.add_argument("--parsed", type=Path, help="parsed directory; defaults to sibling parsed/")
    wiki.add_argument("-o", "--output", required=True, type=Path)

    quartz = subparsers.add_parser(
        "quartz",
        aliases=["publish-quartz"],
        help="publish or resume Quartz output from an existing integrated collection",
    )
    quartz.add_argument("input", type=Path, help="existing collection directory")
    quartz.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Quartz output; defaults to <collection>/quartz",
    )
    quartz.add_argument("--site-title", default="신뢰성 분석 LLM Wiki")
    return parser


def _add_parse_arguments(parser: argparse.ArgumentParser, *, include_semantic: bool) -> None:
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", required=True, type=Path)
    parser.add_argument("--allow-existing-output", action="store_true")
    parser.add_argument("--rendered-slides-dir", type=Path)
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--render-backend", choices=("auto", "powerpoint", "libreoffice"), default="auto")
    parser.add_argument("--source-padding-ratio", type=float, default=0.002)
    parser.add_argument("--model-padding-px", type=int, default=24)
    parser.add_argument(
        "--include-images",
        action="store_true",
        help="opt in to embedded picture extraction (disabled by default)",
    )
    parser.add_argument("--include-empty-shapes", action="store_true")
    parser.add_argument("--strict-extraction", action="store_true")
    parser.add_argument("--strict-ocr", action="store_true")
    parser.add_argument("--office-binary")
    parser.add_argument("--pdf-binary")
    parser.add_argument(
        "--ocr",
        choices=("none", "paddle_cli", "openai_vlm", "command", "paddle_then_vlm", "command_then_vlm"),
        default="none",
    )
    parser.add_argument("--paddle-executable", default="paddleocr")
    parser.add_argument("--paddle-device")
    parser.add_argument("--paddle-engine")
    parser.add_argument(
        "--paddle-extra-arg",
        action="append",
        default=[],
        help="repeat for additional Paddle CLI tokens; use --paddle-extra-arg=--flag for leading dashes",
    )
    parser.add_argument("--ocr-command-json", help='JSON argv, e.g. ["python","infer.py","{image}","{output}"]')
    parser.add_argument("--vlm-base-url")
    parser.add_argument("--vlm-model")
    parser.add_argument("--vlm-api-key-env", default="OPENAI_API_KEY")
    parser.add_argument("--vlm-response-format", choices=("none", "json_object", "json_schema"), default="none")
    if include_semantic:
        parser.add_argument("--synthesize", action="store_true", help="continue through semantic and Wiki stages")
        _add_semantic_arguments(parser, coverage_default="complete")


def _add_semantic_arguments(
    parser: argparse.ArgumentParser,
    *,
    coverage_default: str,
) -> None:
    parser.add_argument("--goal", "--semantic-goal", dest="semantic_goal", default=DEFAULT_SEMANTIC_GOAL)
    parser.add_argument(
        "--coverage-policy",
        choices=("selected", "complete"),
        default=coverage_default,
    )
    parser.add_argument("--language", default="Korean")
    parser.add_argument("--llm-base-url")
    parser.add_argument("--llm-model")
    parser.add_argument("--llm-api-key-env", default="OPENAI_API_KEY")
    parser.add_argument("--max-input-chars", type=int, default=36_000)
    parser.add_argument("--max-output-tokens", type=int, default=4_096)
    parser.add_argument("--max-topics", type=int, default=64)
    parser.add_argument("--repair-attempts", type=int, default=2)
    parser.add_argument("--no-topic-discovery", action="store_true")


def _ensure_output(
    path: Path,
    allow_existing: bool,
    *,
    stages: tuple[str, ...] = (),
) -> None:
    if path.exists() and not path.is_dir():
        raise ValueError(f"output path exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()) and not allow_existing:
        raise ValueError(
            f"output directory is not empty: {path}; choose a new or empty directory"
        )
    for stage in stages:
        target = path / stage
        if target.exists() and not target.is_dir():
            raise ValueError(f"{stage} output exists and is not a directory: {target}")
        if target.exists() and any(target.iterdir()):
            raise ValueError(f"{stage} output directory is not empty: {target}")


def _ocr_adapter(args: argparse.Namespace):
    if args.ocr == "none":
        return None

    def paddle():
        return PaddleOCRCLIAdapter(
            executable=args.paddle_executable,
            device=args.paddle_device,
            engine=args.paddle_engine,
            extra_args=args.paddle_extra_arg,
        )

    def command():
        if not args.ocr_command_json:
            raise ValueError("--ocr command requires --ocr-command-json")
        argv = json.loads(args.ocr_command_json)
        if not isinstance(argv, list) or not all(isinstance(item, str) for item in argv):
            raise ValueError("--ocr-command-json must be a JSON array of strings")
        return CommandOCRAdapter(argv)

    def vlm():
        base_url = args.vlm_base_url or os.getenv("OPENAI_BASE_URL")
        model = args.vlm_model or os.getenv("OPENAI_VLM_MODEL") or os.getenv("OPENAI_MODEL")
        if not base_url or not model:
            raise ValueError("VLM OCR requires --vlm-base-url and --vlm-model (or matching env vars)")
        return OpenAICompatibleVLMAdapter(
            base_url=base_url,
            model=model,
            api_key=os.getenv(args.vlm_api_key_env),
            response_format=args.vlm_response_format,
        )

    if args.ocr == "paddle_cli":
        return paddle()
    if args.ocr == "command":
        return command()
    if args.ocr == "openai_vlm":
        return vlm()
    if args.ocr == "paddle_then_vlm":
        return FallbackOCRAdapter([paddle(), vlm()])
    if args.ocr == "command_then_vlm":
        return FallbackOCRAdapter([command(), vlm()])
    raise AssertionError(args.ocr)


def _semantic_backend(args: argparse.Namespace) -> OpenAICompatibleClient:
    base_url = (
        args.llm_base_url
        or getattr(args, "vlm_base_url", None)
        or os.getenv("OPENAI_BASE_URL")
    )
    model = args.llm_model or os.getenv("OPENAI_LLM_MODEL") or os.getenv("OPENAI_MODEL")
    if not base_url or not model:
        raise ValueError("semantic stage requires --llm-base-url and --llm-model (or matching env vars)")
    return OpenAICompatibleClient(
        base_url=base_url,
        model=model,
        api_key=os.getenv(args.llm_api_key_env),
    )


def _semantic_config(args: argparse.Namespace) -> SemanticConfig:
    return SemanticConfig(
        goal=args.semantic_goal,
        coverage_policy=args.coverage_policy,
        language=args.language,
        max_input_chars=args.max_input_chars,
        max_output_tokens=args.max_output_tokens,
        max_topics=args.max_topics,
        repair_attempts=args.repair_attempts,
        discover_topics=not args.no_topic_discovery,
    )


def _run_parse(args: argparse.Namespace, *, continue_to_wiki: bool) -> int:
    _ensure_output(
        args.output,
        args.allow_existing_output,
        stages=("parsed", "semantic", "wiki"),
    )
    adapter = _ocr_adapter(args)
    backend = _semantic_backend(args) if continue_to_wiki else None
    result = run_pipeline(
        args.input,
        args.output,
        config=PipelineConfig(
            render_backend=args.render_backend,
            dpi=args.dpi,
            source_padding_ratio=args.source_padding_ratio,
            model_padding_px=args.model_padding_px,
            include_images=args.include_images,
            include_empty_shapes=args.include_empty_shapes,
            strict_extraction=args.strict_extraction,
            strict_ocr=args.strict_ocr,
            office_binary=args.office_binary,
            pdf_binary=args.pdf_binary,
        ),
        ocr_adapter=adapter,
        rendered_slides_dir=args.rendered_slides_dir,
        synthesis_backend=backend,
        synthesis_config=_semantic_config(args) if continue_to_wiki else None,
    )
    print(json.dumps(_result_summary(result), ensure_ascii=False, indent=2))
    return 0


def _parse(args: argparse.Namespace) -> int:
    return _run_parse(args, continue_to_wiki=False)


def _run(args: argparse.Namespace) -> int:
    return _run_parse(args, continue_to_wiki=bool(args.synthesize))


def _corpus_dir(value: Path) -> Path:
    root = value.expanduser().resolve()
    if (root / "provenance.jsonl").is_file():
        return root
    if (root / "corpus" / "provenance.jsonl").is_file():
        return root / "corpus"
    raise ValueError(f"parsed provenance not found below: {root}")


def _organize(args: argparse.Namespace) -> int:
    _ensure_output(args.output, False)
    result = build_semantic_output(
        _corpus_dir(args.input),
        backend=_semantic_backend(args),
        output_dir=args.output,
        config=_semantic_config(args),
    )
    print(
        json.dumps(
            {
                "semantic_dir": str(result.output_dir),
                "documents": result.document_count,
                "selected_blocks": len(result.selected_citations),
                "omitted_blocks": len(result.omitted_citations),
                "fallback_documents": len(result.fallback_documents),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def _wiki(args: argparse.Namespace) -> int:
    _ensure_output(args.output, False)
    semantic_dir = args.input.expanduser().resolve()
    parsed_dir = (
        args.parsed.expanduser().resolve()
        if args.parsed is not None
        else semantic_dir.parent / "parsed"
    )
    result = publish_wiki(
        semantic_dir,
        _corpus_dir(parsed_dir),
        args.output,
    )
    print(
        json.dumps(
            {
                "wiki_dir": str(result.output_dir),
                "pages": result.page_count,
                "index": str(result.index_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def _quartz(args: argparse.Namespace) -> int:
    # Preserve lexical components so the publisher can reject symlinks and
    # Windows junctions instead of silently resolving through them.
    collection = args.input.expanduser().absolute()
    output = (
        args.output.expanduser().absolute()
        if args.output is not None
        else collection / "quartz"
    )
    result = publish_quartz(
        collection,
        collection / "integrated",
        output,
        site_title=args.site_title,
    )
    print(
        json.dumps(
            {
                "quartz_dir": str(result.output_dir),
                "content_dir": str(result.content_dir),
                "pages": result.page_count,
                "prs": result.pr_count,
                "entities": result.entity_count,
                "relationships": getattr(result, "relationship_count", 0),
                "assets": len(result.asset_paths),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def _convert(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    source = args.input.expanduser().resolve()
    output = args.output.expanduser().resolve() if args.output else config.output.path_for(source)
    for warning in config.warnings:
        print(f"pptx-wiki: warning: {warning}", file=sys.stderr)
    preflight = {
        "config": str(config.source_path),
        "input": str(source),
        "output": str(output),
        "render": {"backend": config.render.backend, "dpi": config.render.dpi},
        "extraction": {"include_images": config.extraction.include_images},
        "ocr": {
            "enabled": config.ocr.enabled,
            "backend": config.ocr.backend,
            "model": (
                config.ocr.local_model.model
                if "local" in config.ocr.backend
                else config.vlm_api.model if "vlm" in config.ocr.backend else None
            ),
        },
        "semantic": {
            "enabled": config.semantic.enabled,
            "goal": config.semantic.goal if config.semantic.enabled else None,
            "model": config.llm_api.model if config.semantic.enabled else None,
        },
        "wiki": {"enabled": config.wiki.enabled},
    }
    print("Preflight (secrets redacted):")
    print(json.dumps(preflight, ensure_ascii=False, indent=2))
    result = run_configured(source, config, output_override=output)
    print("Completed:")
    print(json.dumps(_result_summary(result), ensure_ascii=False, indent=2))
    return 0


def _collection(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    output = args.output.expanduser().absolute()
    for warning in config.warnings:
        print(f"pptx-wiki: warning: {warning}", file=sys.stderr)
    preflight = {
        "config": str(config.source_path),
        "inputs": [str(value.expanduser().resolve()) for value in args.input],
        "recursive": bool(args.recursive),
        "output": str(output),
        "extraction": {"include_images": config.extraction.include_images},
        "ocr": {
            "enabled": config.ocr.enabled,
            "backend": config.ocr.backend,
            "model": (
                config.ocr.local_model.model
                if "local" in config.ocr.backend
                else config.vlm_api.model if "vlm" in config.ocr.backend else None
            ),
        },
        "semantic": {
            "goal": config.semantic.goal or DEFAULT_COLLECTION_GOAL,
            "coverage_policy": config.semantic.coverage_policy,
            "model": config.llm_api.model,
            "kg_profile": config.semantic.kg_profile,
            "max_relationships": config.semantic.max_relationships,
        },
        "quartz": {"site_title": args.site_title},
    }
    print("Collection preflight (secrets redacted):")
    print(json.dumps(preflight, ensure_ascii=False, indent=2))
    result = run_configured_collection(
        args.input,
        config,
        output_dir=output,
        recursive=bool(args.recursive),
        site_title=args.site_title,
        max_entities=args.max_entities,
        max_files=args.max_files,
        max_total_bytes=args.max_total_mib * 1024 * 1024,
    )
    summary = {
        "output_dir": str(result.output_dir),
        "input_files": result.input_count,
        "unique_decks": result.unique_source_count,
        "pr_numbers": list(result.pr_numbers),
        "parsed_slides": sum(source.parsed.corpus.slide_count for source in result.sources),
        "semantic_markdown_files": len(result.sources),
        "entities": result.integrated.entity_count,
        "relationships": result.integrated.relationship_count,
        "integrated_pages": result.integrated.page_count,
        "quartz_pages": result.quartz.page_count,
        "quartz_content": str(result.quartz.content_dir),
    }
    print("Completed:")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


def _result_summary(result) -> dict[str, object]:
    semantic = getattr(result, "semantic", None)
    wiki = getattr(result, "wiki", None)
    return {
        "output_dir": str(result.output_dir),
        "parsed_dir": str(result.parsed_dir),
        "slides": result.corpus.slide_count,
        "blocks": result.corpus.block_count,
        "ocr_successes": result.ocr_successes,
        "ocr_failures": result.ocr_failures,
        "qa_errors": sum(issue.severity == "error" for issue in result.issues),
        "semantic_documents": semantic.document_count if semantic else 0,
        "semantic_omitted": len(semantic.omitted_citations) if semantic else 0,
        "wiki_pages": (
            getattr(wiki, "page_count", getattr(wiki, "topic_count", 0)) if wiki else 0
        ),
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "convert":
            return _convert(args)
        if args.command in {"batch", "collection"}:
            return _collection(args)
        if args.command == "parse":
            return _parse(args)
        if args.command == "run":
            return _run(args)
        if args.command in {"organize", "semantic"}:
            return _organize(args)
        if args.command == "wiki":
            return _wiki(args)
        if args.command in {"quartz", "publish-quartz"}:
            return _quartz(args)
    except (OSError, RuntimeError, ValueError) as exc:
        parser.exit(2, f"pptx-wiki: error: {exc}\n")
    return 2


if __name__ == "__main__":
    sys.exit(main())
