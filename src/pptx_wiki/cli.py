from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
from typing import Sequence

from .config import load_config
from .configured import run_configured
from .ocr import CommandOCRAdapter, FallbackOCRAdapter, OpenAICompatibleVLMAdapter, PaddleOCRCLIAdapter
from .pipeline import PipelineConfig, run_pipeline
from .synthesis import OpenAICompatibleClient, SynthesisConfig


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-wiki",
        description="Convert PPTX native structure and isolated visual ROIs into a grounded LLM wiki.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    convert = subparsers.add_parser(
        "convert",
        help="convert one PPTX using a trusted YAML config (Windows wrapper mode)",
    )
    convert.add_argument("input", type=Path)
    convert.add_argument("--config", required=True, type=Path)
    convert.add_argument("-o", "--output", type=Path)

    run = subparsers.add_parser("run", help="run native extraction, optional OCR, and optional synthesis")
    run.add_argument("input", type=Path)
    run.add_argument("-o", "--output", required=True, type=Path)
    run.add_argument("--allow-existing-output", action="store_true")
    run.add_argument("--rendered-slides-dir", type=Path)
    run.add_argument("--dpi", type=int, default=300)
    run.add_argument("--render-backend", choices=("auto", "powerpoint", "libreoffice"), default="auto")
    run.add_argument("--source-padding-ratio", type=float, default=0.002)
    run.add_argument("--model-padding-px", type=int, default=24)
    run.add_argument("--include-empty-shapes", action="store_true")
    run.add_argument("--strict-extraction", action="store_true")
    run.add_argument("--strict-ocr", action="store_true")
    run.add_argument("--office-binary")
    run.add_argument("--pdf-binary")

    run.add_argument(
        "--ocr",
        choices=("none", "paddle_cli", "openai_vlm", "command", "paddle_then_vlm", "command_then_vlm"),
        default="none",
    )
    run.add_argument("--paddle-executable", default="paddleocr")
    run.add_argument("--paddle-device")
    run.add_argument("--paddle-engine")
    run.add_argument(
        "--paddle-extra-arg",
        action="append",
        default=[],
        help="repeat for additional Paddle CLI tokens; use --paddle-extra-arg=--flag for leading dashes",
    )
    run.add_argument("--ocr-command-json", help='JSON argv, e.g. ["python","infer.py","{image}","{output}"]')
    run.add_argument("--vlm-base-url")
    run.add_argument("--vlm-model")
    run.add_argument("--vlm-api-key-env", default="OPENAI_API_KEY")
    run.add_argument("--vlm-response-format", choices=("none", "json_object", "json_schema"), default="none")

    run.add_argument("--synthesize", action="store_true")
    run.add_argument("--llm-base-url")
    run.add_argument("--llm-model")
    run.add_argument("--llm-api-key-env", default="OPENAI_API_KEY")
    run.add_argument("--max-input-chars", type=int, default=36_000)
    run.add_argument("--max-output-tokens", type=int, default=4_096)
    run.add_argument("--no-topic-discovery", action="store_true")
    return parser


def _ensure_output(path: Path, allow_existing: bool) -> None:
    if path.exists() and any(path.iterdir()) and not allow_existing:
        raise ValueError(
            f"output directory is not empty: {path}; choose a new directory or pass --allow-existing-output"
        )


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


def _synthesis_backend(args: argparse.Namespace):
    if not args.synthesize:
        return None
    base_url = args.llm_base_url or args.vlm_base_url or os.getenv("OPENAI_BASE_URL")
    model = args.llm_model or os.getenv("OPENAI_LLM_MODEL") or os.getenv("OPENAI_MODEL")
    if not base_url or not model:
        raise ValueError("--synthesize requires --llm-base-url and --llm-model (or matching env vars)")
    return OpenAICompatibleClient(
        base_url=base_url,
        model=model,
        api_key=os.getenv(args.llm_api_key_env),
    )


def _run(args: argparse.Namespace) -> int:
    _ensure_output(args.output, args.allow_existing_output)
    adapter = _ocr_adapter(args)
    backend = _synthesis_backend(args)
    result = run_pipeline(
        args.input,
        args.output,
        config=PipelineConfig(
            render_backend=args.render_backend,
            dpi=args.dpi,
            source_padding_ratio=args.source_padding_ratio,
            model_padding_px=args.model_padding_px,
            include_empty_shapes=args.include_empty_shapes,
            strict_extraction=args.strict_extraction,
            strict_ocr=args.strict_ocr,
            office_binary=args.office_binary,
            pdf_binary=args.pdf_binary,
        ),
        ocr_adapter=adapter,
        rendered_slides_dir=args.rendered_slides_dir,
        synthesis_backend=backend,
        synthesis_config=SynthesisConfig(
            max_input_chars=args.max_input_chars,
            max_output_tokens=args.max_output_tokens,
            discover_topics=not args.no_topic_discovery,
        ),
    )
    print(json.dumps(_result_summary(result), ensure_ascii=False, indent=2))
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
        "ocr": {
            "enabled": config.ocr.enabled,
            "backend": config.ocr.backend,
            "model": (
                config.ocr.local_model.model
                if "local" in config.ocr.backend
                else config.vlm_api.model if "vlm" in config.ocr.backend else None
            ),
        },
        "wiki": {
            "enabled": config.wiki.enabled,
            "model": config.llm_api.model if config.wiki.enabled else None,
        },
    }
    print("Preflight (secrets redacted):")
    print(json.dumps(preflight, ensure_ascii=False, indent=2))
    result = run_configured(source, config, output_override=output)
    print("Completed:")
    print(json.dumps(_result_summary(result), ensure_ascii=False, indent=2))
    return 0


def _result_summary(result) -> dict[str, object]:
    return {
        "output_dir": str(result.output_dir),
        "slides": result.corpus.slide_count,
        "blocks": result.corpus.block_count,
        "ocr_successes": result.ocr_successes,
        "ocr_failures": result.ocr_failures,
        "qa_errors": sum(issue.severity == "error" for issue in result.issues),
        "wiki_topics": result.wiki.topic_count if result.wiki else 0,
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "convert":
            return _convert(args)
        if args.command == "run":
            return _run(args)
    except (OSError, RuntimeError, ValueError) as exc:
        parser.exit(2, f"pptx-wiki: error: {exc}\n")
    return 2


if __name__ == "__main__":
    sys.exit(main())
