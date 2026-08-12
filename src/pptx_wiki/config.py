from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import ipaddress
import math
import os
from pathlib import Path
import re
from typing import Any, Literal, Mapping
from urllib.parse import urlsplit

import yaml
from yaml.constructor import ConstructorError
from yaml.events import AliasEvent


RenderBackend = Literal["auto", "powerpoint", "libreoffice"]
CoveragePolicy = Literal["selected", "complete"]
OCRBackend = Literal[
    "none",
    "openai_vlm",
    "local_model",
    "paddle_cli",
    "command",
    "local_then_vlm",
    "paddle_then_vlm",
    "command_then_vlm",
]
LocalModelName = Literal[
    "paddleocr_vl_16",
    "paddleocr_vl_15",
    "monkeyocr_v2_b",
    "ovisocr2",
]
_ENV_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_LOCAL_DEVICE = re.compile(r"^(?:auto|cpu|cuda(?::[0-9]+)?|gpu(?::[0-9]+)?)$")
_WINDOWS_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


class _StrictSafeLoader(yaml.SafeLoader):
    """SafeLoader that also rejects aliases, anchors, merges and duplicates."""

    def compose_node(self, parent: Any, index: Any) -> Any:
        if self.check_event(AliasEvent):
            event = self.peek_event()
            raise ConstructorError(None, None, "YAML aliases are not allowed", event.start_mark)
        event = self.peek_event()
        if getattr(event, "anchor", None) is not None:
            raise ConstructorError(None, None, "YAML anchors are not allowed", event.start_mark)
        return super().compose_node(parent, index)

    def construct_mapping(self, node: Any, deep: bool = False) -> dict[Any, Any]:
        seen: set[Any] = set()
        for key_node, _ in node.value:
            if key_node.tag == "tag:yaml.org,2002:merge" or getattr(key_node, "value", None) == "<<":
                raise ConstructorError(None, None, "YAML merge keys are not allowed", key_node.start_mark)
            key = self.construct_object(key_node, deep=deep)
            if key in seen:
                raise ConstructorError(None, None, f"duplicate YAML key: {key!r}", key_node.start_mark)
            seen.add(key)
        return super().construct_mapping(node, deep=deep)


@dataclass(frozen=True, slots=True)
class OutputSettings:
    directory: Path
    subdirectory_per_pptx: bool = True
    naming: Literal["stem", "stem_hash"] = "stem"
    allow_existing: bool = False

    def path_for(self, pptx_path: str | Path) -> Path:
        if not self.subdirectory_per_pptx:
            return self.directory
        source = Path(pptx_path)
        stem = _safe_output_stem(source.stem)
        if self.naming == "stem_hash":
            stem = f"{stem}-{_file_sha256(source)[:10]}"
        return self.directory / stem


@dataclass(frozen=True, slots=True)
class RenderSettings:
    backend: RenderBackend = "powerpoint"
    dpi: int = 300
    source_padding_ratio: float = 0.002
    model_padding_px: int = 24
    block_external_resources: bool = True
    rendered_slides_dir: Path | None = None
    office_binary: str | None = None
    pdf_binary: str | None = None


@dataclass(frozen=True, slots=True)
class ExtractionSettings:
    include_images: bool = False
    include_empty_shapes: bool = False
    strict: bool = False


@dataclass(frozen=True, slots=True)
class APISettings:
    base_url: str = ""
    model: str = ""
    api_key: str = field(default="", repr=False)
    api_key_env: str = ""
    timeout_seconds: float = 180.0
    max_tokens: int = 8192

    def resolved_api_key(self) -> str | None:
        if self.api_key:
            return self.api_key
        if self.api_key_env:
            return os.getenv(self.api_key_env) or None
        return None


@dataclass(frozen=True, slots=True)
class PaddleSettings:
    executable: str = "paddleocr"
    pipeline_version: str = "v1.6"
    device: str | None = None
    engine: str | None = None
    timeout_seconds: float = 300.0
    extra_args: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class CommandSettings:
    argv: tuple[str, ...] = ()
    cwd: Path | None = None
    timeout_seconds: float = 300.0


@dataclass(frozen=True, slots=True)
class LocalModelSettings:
    """One of the bundled, version-pinned local OCR workers.

    Every profile owns a separate virtual environment under
    ``workers_directory/<model>/.venv`` and a separate downloaded snapshot under
    ``models_directory/<model>``.  This deliberately prevents Paddle, vLLM and
    Transformers dependency sets from contaminating the main application.
    """

    model: LocalModelName = "paddleocr_vl_16"
    workers_directory: Path = Path("workers")
    models_directory: Path = Path("models")
    python_executable: Path | None = None
    worker_script: Path | None = None
    device: str = "auto"
    dtype: Literal["auto", "float32", "float16", "bfloat16"] = "auto"
    startup_timeout_seconds: float = 900.0
    request_timeout_seconds: float = 600.0
    max_new_tokens: int = 16_384
    fallback_to_vlm: bool = False

    @property
    def profile_directory(self) -> Path:
        return self.workers_directory / self.model

    @property
    def resolved_worker_script(self) -> Path:
        return self.worker_script or self.profile_directory / "worker.py"

    @property
    def resolved_model_directory(self) -> Path:
        return self.models_directory / self.model

    def resolved_python_executable(self, *, windows: bool | None = None) -> Path:
        if self.python_executable is not None:
            return self.python_executable
        is_windows = os.name == "nt" if windows is None else windows
        relative = Path("Scripts/python.exe") if is_windows else Path("bin/python")
        return self.profile_directory / ".venv" / relative


@dataclass(frozen=True, slots=True)
class OCRSettings:
    enabled: bool = True
    backend: OCRBackend = "openai_vlm"
    strict: bool = False
    response_format: Literal["none", "json_object", "json_schema"] = "json_schema"
    image_detail: Literal["auto", "low", "high"] = "high"
    local_model: LocalModelSettings = field(default_factory=LocalModelSettings)
    paddle: PaddleSettings = field(default_factory=PaddleSettings)
    command: CommandSettings = field(default_factory=CommandSettings)


@dataclass(frozen=True, slots=True)
class SemanticSettings:
    enabled: bool = True
    goal: str = ""
    coverage_policy: CoveragePolicy = "complete"
    language: str = "Korean"
    max_input_chars: int = 36_000
    max_output_tokens: int = 4_096
    max_topics: int = 64
    repair_attempts: int = 2
    discover_topics: bool = True
    kg_profile: Literal["none", "semiconductor_reliability"] = (
        "semiconductor_reliability"
    )
    max_relationships: int = 512


@dataclass(frozen=True, slots=True)
class WikiSettings:
    enabled: bool = True


@dataclass(frozen=True, slots=True)
class NetworkSettings:
    allow_remote_endpoints: bool = False
    allow_insecure_http: bool = False


@dataclass(frozen=True, slots=True)
class AppConfig:
    source_path: Path
    output: OutputSettings
    render: RenderSettings
    extraction: ExtractionSettings
    vlm_api: APISettings
    llm_api: APISettings
    ocr: OCRSettings
    semantic: SemanticSettings
    wiki: WikiSettings
    network: NetworkSettings
    warnings: tuple[str, ...] = ()


def load_config(path: str | Path) -> AppConfig:
    """Load a strictly validated YAML configuration using ``safe_load``."""

    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.stat().st_size > 256 * 1024:
        raise ValueError("config file exceeds the 256 KiB size limit")
    try:
        value = yaml.load(source.read_text(encoding="utf-8-sig"), Loader=_StrictSafeLoader)
    except yaml.YAMLError as exc:
        raise ValueError(f"invalid YAML in {source}: {exc}") from exc
    root = _mapping(value or {}, "config")
    version = _integer(root.get("version", 1), "version", minimum=1)
    if version not in {1, 2}:
        raise ValueError(f"unsupported config version: {version}")
    allowed_root_keys = {
        "version",
        "output",
        "render",
        "extraction",
        "vlm_api",
        "llm_api",
        "ocr",
        "wiki",
        "network",
    }
    if version == 2:
        allowed_root_keys.add("semantic")
    _check_keys(root, allowed_root_keys, "config")
    base = source.parent
    warnings: list[str] = []

    output_value = _section(root, "output")
    _check_keys(
        output_value,
        {"directory", "subdirectory_per_pptx", "naming", "allow_existing"},
        "output",
    )
    output_naming = _choice(
        output_value.get("naming", "stem"), {"stem", "stem_hash"}, "output.naming"
    )
    output = OutputSettings(
        directory=_path(output_value.get("directory", "./output"), base, "output.directory"),
        subdirectory_per_pptx=_boolean(
            output_value.get("subdirectory_per_pptx", True), "output.subdirectory_per_pptx"
        ),
        naming=output_naming,  # type: ignore[arg-type]
        allow_existing=_boolean(output_value.get("allow_existing", False), "output.allow_existing"),
    )

    render_value = _section(root, "render")
    _check_keys(
        render_value,
        {
            "backend",
            "dpi",
            "source_padding_ratio",
            "model_padding_px",
            "block_external_resources",
            "rendered_slides_dir",
            "office_binary",
            "pdf_binary",
        },
        "render",
    )
    render_backend = _choice(
        render_value.get("backend", "powerpoint"),
        {"auto", "powerpoint", "libreoffice"},
        "render.backend",
    )
    rendered_dir_value = render_value.get("rendered_slides_dir")
    render = RenderSettings(
        backend=render_backend,  # type: ignore[arg-type]
        dpi=_integer(render_value.get("dpi", 300), "render.dpi", minimum=72, maximum=600),
        source_padding_ratio=_number(
            render_value.get("source_padding_ratio", 0.002),
            "render.source_padding_ratio",
            minimum=0.0,
            maximum=0.1,
        ),
        model_padding_px=_integer(
            render_value.get("model_padding_px", 24),
            "render.model_padding_px",
            minimum=0,
            maximum=512,
        ),
        block_external_resources=_boolean(
            render_value.get("block_external_resources", True),
            "render.block_external_resources",
        ),
        rendered_slides_dir=(
            _path(rendered_dir_value, base, "render.rendered_slides_dir")
            if rendered_dir_value not in {None, ""}
            else None
        ),
        office_binary=_optional_string(render_value.get("office_binary"), "render.office_binary"),
        pdf_binary=_optional_string(render_value.get("pdf_binary"), "render.pdf_binary"),
    )

    extraction_value = _section(root, "extraction")
    _check_keys(
        extraction_value,
        {"include_images", "include_empty_shapes", "strict"},
        "extraction",
    )
    extraction = ExtractionSettings(
        include_images=_boolean(
            extraction_value.get("include_images", False),
            "extraction.include_images",
        ),
        include_empty_shapes=_boolean(
            extraction_value.get("include_empty_shapes", False), "extraction.include_empty_shapes"
        ),
        strict=_boolean(extraction_value.get("strict", False), "extraction.strict"),
    )

    vlm_api = _api_settings(_section(root, "vlm_api"), "vlm_api")
    llm_api = _api_settings(_section(root, "llm_api"), "llm_api")
    if vlm_api.api_key:
        warnings.append("vlm_api.api_key is stored as plaintext; api_key_env is safer")
    if llm_api.api_key:
        warnings.append("llm_api.api_key is stored as plaintext; api_key_env is safer")

    ocr_value = _section(root, "ocr")
    _check_keys(
        ocr_value,
        {
            "enabled",
            "backend",
            "strict",
            "response_format",
            "image_detail",
            "local_model",
            "paddle",
            "command",
        },
        "ocr",
    )
    local_model_value = _section(ocr_value, "local_model")
    _check_keys(
        local_model_value,
        {
            "model",
            "workers_directory",
            "models_directory",
            "python_executable",
            "worker_script",
            "device",
            "dtype",
            "startup_timeout_seconds",
            "request_timeout_seconds",
            "max_new_tokens",
            "fallback_to_vlm",
        },
        "ocr.local_model",
    )
    paddle_value = _section(ocr_value, "paddle")
    _check_keys(
        paddle_value,
        {"executable", "pipeline_version", "device", "engine", "timeout_seconds", "extra_args"},
        "ocr.paddle",
    )
    command_value = _section(ocr_value, "command")
    _check_keys(command_value, {"argv", "cwd", "timeout_seconds"}, "ocr.command")
    ocr_backend = _choice(
        ocr_value.get("backend", "openai_vlm"),
        {
            "none",
            "openai_vlm",
            "local_model",
            "paddle_cli",
            "command",
            "local_then_vlm",
            "paddle_then_vlm",
            "command_then_vlm",
        },
        "ocr.backend",
    )
    response_format = _choice(
        ocr_value.get("response_format", "json_schema"),
        {"none", "json_object", "json_schema"},
        "ocr.response_format",
    )
    image_detail = _choice(
        ocr_value.get("image_detail", "high"), {"auto", "low", "high"}, "ocr.image_detail"
    )
    command_cwd = command_value.get("cwd")
    local_python = local_model_value.get("python_executable")
    local_worker = local_model_value.get("worker_script")
    local_model_name = _choice(
        local_model_value.get("model", "paddleocr_vl_16"),
        {"paddleocr_vl_16", "paddleocr_vl_15", "monkeyocr_v2_b", "ovisocr2"},
        "ocr.local_model.model",
    )
    local_dtype = _choice(
        local_model_value.get("dtype", "auto"),
        {"auto", "float32", "float16", "bfloat16"},
        "ocr.local_model.dtype",
    )
    ocr = OCRSettings(
        enabled=_boolean(ocr_value.get("enabled", True), "ocr.enabled"),
        backend=ocr_backend,  # type: ignore[arg-type]
        strict=_boolean(ocr_value.get("strict", False), "ocr.strict"),
        response_format=response_format,  # type: ignore[arg-type]
        image_detail=image_detail,  # type: ignore[arg-type]
        local_model=LocalModelSettings(
            model=local_model_name,  # type: ignore[arg-type]
            workers_directory=_path(
                local_model_value.get("workers_directory", "./workers"),
                base,
                "ocr.local_model.workers_directory",
            ),
            models_directory=_path(
                local_model_value.get("models_directory", "./models"),
                base,
                "ocr.local_model.models_directory",
            ),
            python_executable=(
                _path(local_python, base, "ocr.local_model.python_executable")
                if local_python not in {None, ""}
                else None
            ),
            worker_script=(
                _path(local_worker, base, "ocr.local_model.worker_script")
                if local_worker not in {None, ""}
                else None
            ),
            device=_local_device(
                local_model_value.get("device", "auto"), "ocr.local_model.device"
            ),
            dtype=local_dtype,  # type: ignore[arg-type]
            startup_timeout_seconds=_number(
                local_model_value.get("startup_timeout_seconds", 900),
                "ocr.local_model.startup_timeout_seconds",
                minimum=1,
                maximum=7200,
            ),
            request_timeout_seconds=_number(
                local_model_value.get("request_timeout_seconds", 600),
                "ocr.local_model.request_timeout_seconds",
                minimum=1,
                maximum=7200,
            ),
            max_new_tokens=_integer(
                local_model_value.get("max_new_tokens", 16_384),
                "ocr.local_model.max_new_tokens",
                minimum=64,
                maximum=65_536,
            ),
            fallback_to_vlm=_boolean(
                local_model_value.get("fallback_to_vlm", False),
                "ocr.local_model.fallback_to_vlm",
            ),
        ),
        paddle=PaddleSettings(
            executable=_string(paddle_value.get("executable", "paddleocr"), "ocr.paddle.executable"),
            pipeline_version=_string(
                paddle_value.get("pipeline_version", "v1.6"), "ocr.paddle.pipeline_version"
            ),
            device=_optional_string(paddle_value.get("device"), "ocr.paddle.device"),
            engine=_optional_string(paddle_value.get("engine"), "ocr.paddle.engine"),
            timeout_seconds=_number(
                paddle_value.get("timeout_seconds", 300),
                "ocr.paddle.timeout_seconds",
                minimum=1,
                maximum=3600,
            ),
            extra_args=_string_tuple(paddle_value.get("extra_args", []), "ocr.paddle.extra_args"),
        ),
        command=CommandSettings(
            argv=_string_tuple(command_value.get("argv", []), "ocr.command.argv"),
            cwd=_path(command_cwd, base, "ocr.command.cwd") if command_cwd not in {None, ""} else None,
            timeout_seconds=_number(
                command_value.get("timeout_seconds", 300),
                "ocr.command.timeout_seconds",
                minimum=1,
                maximum=3600,
            ),
        ),
    )

    wiki_value = _section(root, "wiki")
    if version == 1:
        _check_keys(
            wiki_value,
            {
                "enabled",
                "language",
                "max_input_chars",
                "max_output_tokens",
                "max_topics",
                "repair_attempts",
                "discover_topics",
                "kg_profile",
                "max_relationships",
            },
            "wiki",
        )
        legacy_enabled = _boolean(wiki_value.get("enabled", True), "wiki.enabled")
        semantic = _semantic_settings(
            wiki_value,
            label="wiki",
            enabled=legacy_enabled,
        )
        wiki = WikiSettings(enabled=legacy_enabled)
        warnings.append(
            "config version 1 is deprecated: wiki.enabled is mapped to both "
            "semantic.enabled and wiki.enabled; other wiki settings are mapped to semantic"
        )
    else:
        semantic_value = _section(root, "semantic")
        _check_keys(
            semantic_value,
            {
                "enabled",
                "goal",
                "coverage_policy",
                "language",
                "max_input_chars",
                "max_output_tokens",
                "max_topics",
                "repair_attempts",
                "discover_topics",
                "kg_profile",
                "max_relationships",
            },
            "semantic",
        )
        semantic = _semantic_settings(semantic_value, label="semantic")
        _check_keys(wiki_value, {"enabled"}, "wiki")
        wiki = WikiSettings(
            enabled=_boolean(wiki_value.get("enabled", True), "wiki.enabled")
        )

    network_value = _section(root, "network")
    _check_keys(
        network_value,
        {"allow_remote_endpoints", "allow_insecure_http"},
        "network",
    )
    network = NetworkSettings(
        allow_remote_endpoints=_boolean(
            network_value.get("allow_remote_endpoints", False),
            "network.allow_remote_endpoints",
        ),
        allow_insecure_http=_boolean(
            network_value.get("allow_insecure_http", False),
            "network.allow_insecure_http",
        ),
    )

    _validate_cross_section(vlm_api, llm_api, ocr, semantic, wiki, network)
    return AppConfig(
        source_path=source,
        output=output,
        render=render,
        extraction=extraction,
        vlm_api=vlm_api,
        llm_api=llm_api,
        ocr=ocr,
        semantic=semantic,
        wiki=wiki,
        network=network,
        warnings=tuple(warnings),
    )


def _semantic_settings(
    value: Mapping[str, Any],
    *,
    label: str,
    enabled: bool | None = None,
) -> SemanticSettings:
    return SemanticSettings(
        enabled=(
            _boolean(value.get("enabled", True), f"{label}.enabled")
            if enabled is None
            else enabled
        ),
        goal=_string(value.get("goal", ""), f"{label}.goal"),
        coverage_policy=_choice(
            value.get("coverage_policy", "complete"),
            {"selected", "complete"},
            f"{label}.coverage_policy",
        ),  # type: ignore[arg-type]
        language=_string(value.get("language", "Korean"), f"{label}.language"),
        max_input_chars=_integer(
            value.get("max_input_chars", 36_000),
            f"{label}.max_input_chars",
            minimum=2_000,
            maximum=2_000_000,
        ),
        max_output_tokens=_integer(
            value.get("max_output_tokens", 4_096),
            f"{label}.max_output_tokens",
            minimum=256,
            maximum=131_072,
        ),
        max_topics=_integer(
            value.get("max_topics", 64),
            f"{label}.max_topics",
            minimum=1,
            maximum=512,
        ),
        repair_attempts=_integer(
            value.get("repair_attempts", 2),
            f"{label}.repair_attempts",
            minimum=0,
            maximum=10,
        ),
        discover_topics=_boolean(
            value.get("discover_topics", True), f"{label}.discover_topics"
        ),
        kg_profile=_choice(
            value.get("kg_profile", "semiconductor_reliability"),
            {"none", "semiconductor_reliability"},
            f"{label}.kg_profile",
        ),  # type: ignore[arg-type]
        max_relationships=_integer(
            value.get("max_relationships", 512),
            f"{label}.max_relationships",
            minimum=1,
            maximum=4096,
        ),
    )


def _api_settings(value: Mapping[str, Any], label: str) -> APISettings:
    _check_keys(
        value,
        {"base_url", "model", "api_key", "api_key_env", "timeout_seconds", "max_tokens"},
        label,
    )
    api_key = _string(value.get("api_key", ""), f"{label}.api_key")
    api_key_env = _string(value.get("api_key_env", ""), f"{label}.api_key_env")
    if api_key and api_key_env:
        raise ValueError(f"{label}.api_key and {label}.api_key_env are mutually exclusive")
    if api_key_env and not _ENV_NAME.fullmatch(api_key_env):
        raise ValueError(f"{label}.api_key_env is not a valid environment variable name")
    return APISettings(
        base_url=_string(value.get("base_url", ""), f"{label}.base_url").rstrip("/"),
        model=_string(value.get("model", ""), f"{label}.model"),
        api_key=api_key,
        api_key_env=api_key_env,
        timeout_seconds=_number(
            value.get("timeout_seconds", 180),
            f"{label}.timeout_seconds",
            minimum=1,
            maximum=3600,
        ),
        max_tokens=_integer(
            value.get("max_tokens", 8192),
            f"{label}.max_tokens",
            minimum=256,
            maximum=262_144,
        ),
    )


def _validate_cross_section(
    vlm_api: APISettings,
    llm_api: APISettings,
    ocr: OCRSettings,
    semantic: SemanticSettings,
    wiki: WikiSettings,
    network: NetworkSettings,
) -> None:
    if (
        ocr.local_model.model in {"monkeyocr_v2_b", "ovisocr2"}
        and ocr.local_model.max_new_tokens > 32_768
    ):
        raise ValueError(
            "ocr.local_model.max_new_tokens must be at most 32768 for "
            f"{ocr.local_model.model}"
        )
    local_fallback = ocr.backend == "local_model" and ocr.local_model.fallback_to_vlm
    if ocr.enabled and ocr.backend != "none" and ("vlm" in ocr.backend or local_fallback):
        _require_endpoint(vlm_api, "vlm_api", network)
        _require_secret(vlm_api, "vlm_api")
    if ocr.enabled and "command" in ocr.backend and not ocr.command.argv:
        raise ValueError("ocr.command.argv is required for a command backend")
    if ocr.enabled and ocr.backend in {"local_model", "local_then_vlm"}:
        if not ocr.local_model.device:
            raise ValueError("ocr.local_model.device cannot be empty")
    if wiki.enabled and not semantic.enabled:
        raise ValueError("wiki.enabled requires semantic.enabled in one-shot configuration")
    if semantic.enabled:
        _require_endpoint(llm_api, "llm_api", network)
        _require_secret(llm_api, "llm_api")


def _require_endpoint(endpoint: APISettings, label: str, network: NetworkSettings) -> None:
    parsed = urlsplit(endpoint.base_url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError(f"{label}.base_url must be an HTTP(S) URL")
    if parsed.username or parsed.password:
        raise ValueError(f"{label}.base_url must not contain credentials")
    if parsed.query or parsed.fragment:
        raise ValueError(f"{label}.base_url must not contain a query or fragment")
    hostname = parsed.hostname.casefold()
    loopback = hostname == "localhost"
    try:
        address = ipaddress.ip_address(hostname)
    except ValueError:
        address = None
    if address is not None:
        if address.is_unspecified or address.is_link_local:
            raise ValueError(f"{label}.base_url uses an unsafe bind/link-local address")
        loopback = address.is_loopback
    if not loopback and not network.allow_remote_endpoints:
        raise ValueError(
            f"{label}.base_url is remote; set network.allow_remote_endpoints=true only if document upload is intended"
        )
    if parsed.scheme == "http" and not loopback and not network.allow_insecure_http:
        raise ValueError(
            f"{label}.base_url uses remote plaintext HTTP; use HTTPS or explicitly set network.allow_insecure_http=true"
        )
    if not endpoint.model:
        raise ValueError(f"{label}.model is required")


def _require_secret(endpoint: APISettings, label: str) -> None:
    if endpoint.api_key_env and not os.getenv(endpoint.api_key_env):
        raise ValueError(
            f"environment variable {endpoint.api_key_env!r} configured by {label}.api_key_env is not set; "
            "set it or use an empty api_key_env for an unauthenticated local endpoint"
        )


def _section(value: Mapping[str, Any], name: str) -> Mapping[str, Any]:
    return _mapping(value.get(name, {}), name)


def _mapping(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValueError(f"{label} must be a YAML mapping")
    return value


def _check_keys(value: Mapping[str, Any], allowed: set[str], label: str) -> None:
    unknown = sorted(str(key) for key in value if key not in allowed)
    if unknown:
        raise ValueError(f"unknown {label} setting(s): {', '.join(unknown)}")


def _string(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be a string")
    return value.strip()


def _optional_string(value: Any, label: str) -> str | None:
    if value in {None, ""}:
        return None
    return _string(value, label)


def _local_device(value: Any, label: str) -> str:
    device = _string(value, label).casefold()
    if not _LOCAL_DEVICE.fullmatch(device):
        raise ValueError(
            f"{label} must be auto, cpu, cuda, cuda:N, gpu, or gpu:N"
        )
    return device


def _boolean(value: Any, label: str) -> bool:
    if not isinstance(value, bool):
        raise ValueError(f"{label} must be true or false")
    return value


def _integer(value: Any, label: str, *, minimum: int, maximum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError(f"{label} must be an integer")
    if value < minimum:
        raise ValueError(f"{label} must be at least {minimum}")
    if maximum is not None and value > maximum:
        raise ValueError(f"{label} must be at most {maximum}")
    return value


def _number(
    value: Any,
    label: str,
    *,
    minimum: float,
    maximum: float | None = None,
) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label} must be a number")
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"{label} must be finite")
    if number < minimum or (maximum is not None and number > maximum):
        limit = f" between {minimum} and {maximum}" if maximum is not None else f" at least {minimum}"
        raise ValueError(f"{label} must be{limit}")
    return number


def _choice(value: Any, choices: set[str], label: str) -> str:
    text = _string(value, label)
    if text not in choices:
        raise ValueError(f"{label} must be one of: {', '.join(sorted(choices))}")
    return text


def _path(value: Any, base: Path, label: str) -> Path:
    text = _string(value, label)
    if not text:
        raise ValueError(f"{label} cannot be empty")
    path = Path(os.path.expandvars(text)).expanduser()
    return (base / path).resolve() if not path.is_absolute() else path.resolve()


def _string_tuple(value: Any, label: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{label} must be a YAML list of strings")
    return tuple(value)


def _safe_output_stem(value: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value).strip(" .")
    clean = re.sub(r"\s+", " ", clean)
    if not clean:
        clean = "deck"
    if clean.upper() in _WINDOWS_RESERVED:
        clean = f"_{clean}"
    return clean[:100].rstrip(" .") or "deck"


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


__all__ = [
    "APISettings",
    "AppConfig",
    "CommandSettings",
    "CoveragePolicy",
    "ExtractionSettings",
    "LocalModelName",
    "LocalModelSettings",
    "OCRSettings",
    "NetworkSettings",
    "OutputSettings",
    "PaddleSettings",
    "RenderSettings",
    "SemanticSettings",
    "WikiSettings",
    "load_config",
]
