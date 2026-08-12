from __future__ import annotations

import os
from pathlib import Path

import pytest

import pptx_wiki.configured as configured_module
from pptx_wiki.config import APISettings, OutputSettings, load_config
from pptx_wiki.configured import build_ocr_adapter
from pptx_wiki.ocr import OpenAICompatibleVLMAdapter, PersistentOCRWorkerAdapter


def _base_yaml(extra: str = "") -> str:
    return (
        """version: 1
output:
  directory: ./results
render:
  backend: powerpoint
extraction: {}
vlm_api: {}
llm_api: {}
ocr:
  enabled: false
  backend: none
wiki:
  enabled: false
network: {}
"""
        + extra
    )


def _v2_yaml(*, semantic_enabled: bool, wiki_enabled: bool, llm_api: str = "{}") -> str:
    return f"""version: 2
output:
  directory: ./results
render:
  backend: powerpoint
extraction: {{}}
vlm_api: {{}}
llm_api: {llm_api}
ocr:
  enabled: false
  backend: none
semantic:
  enabled: {str(semantic_enabled).lower()}
wiki:
  enabled: {str(wiki_enabled).lower()}
network: {{}}
"""


def _write(tmp_path: Path, value: str) -> Path:
    path = tmp_path / "config.yml"
    path.write_text(value, encoding="utf-8")
    return path


def test_load_config_resolves_paths_from_config_directory(tmp_path: Path) -> None:
    config = load_config(_write(tmp_path, _base_yaml()))
    assert config.output.directory == (tmp_path / "results").resolve()
    assert config.render.backend == "powerpoint"
    assert config.ocr.enabled is False


def test_version_1_wiki_settings_map_to_semantic_and_emit_warning(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "llm_api: {}",
        "llm_api:\n  base_url: http://127.0.0.1:8001/v1\n  model: local-text",
    ).replace(
        "wiki:\n  enabled: false",
        """wiki:
  enabled: true
  language: English
  max_input_chars: 12000
  max_output_tokens: 2048
  max_topics: 12
  repair_attempts: 1
  discover_topics: false""",
    )

    config = load_config(_write(tmp_path, value))

    assert config.semantic.enabled is True
    assert config.semantic.goal == ""
    assert config.semantic.coverage_policy == "complete"
    assert config.semantic.language == "English"
    assert config.semantic.max_input_chars == 12_000
    assert config.semantic.max_output_tokens == 2_048
    assert config.semantic.max_topics == 12
    assert config.semantic.repair_attempts == 1
    assert config.semantic.discover_topics is False
    assert config.wiki.enabled is True
    assert any("version 1 is deprecated" in warning for warning in config.warnings)


def test_version_2_loads_semantic_settings_separately_from_wiki(tmp_path: Path) -> None:
    value = _v2_yaml(
        semantic_enabled=True,
        wiki_enabled=True,
    ).replace(
        "llm_api: {}\n",
        "llm_api:\n  base_url: http://127.0.0.1:8001/v1\n  model: local-text\n",
    ).replace(
        "semantic:\n  enabled: true",
        """semantic:
  enabled: true
  goal: Focus on deployment
  coverage_policy: selected
  language: English
  max_input_chars: 16000
  max_output_tokens: 3072
  max_topics: 20
  repair_attempts: 3
  discover_topics: false
  kg_profile: none
  max_relationships: 128""",
    )

    config = load_config(_write(tmp_path, value))

    assert config.semantic.goal == "Focus on deployment"
    assert config.semantic.coverage_policy == "selected"
    assert config.semantic.language == "English"
    assert config.semantic.max_input_chars == 16_000
    assert config.semantic.max_output_tokens == 3_072
    assert config.semantic.max_topics == 20
    assert config.semantic.repair_attempts == 3
    assert config.semantic.discover_topics is False
    assert config.semantic.kg_profile == "none"
    assert config.semantic.max_relationships == 128
    assert config.wiki.enabled is True
    assert config.warnings == ()


@pytest.mark.parametrize(
    "semantic_enabled,wiki_enabled,match",
    [
        (True, False, "llm_api.base_url"),
        (False, True, "wiki.enabled requires semantic.enabled"),
    ],
)
def test_version_2_validates_stage_dependencies(
    tmp_path: Path,
    semantic_enabled: bool,
    wiki_enabled: bool,
    match: str,
) -> None:
    with pytest.raises(ValueError, match=match):
        load_config(
            _write(
                tmp_path,
                _v2_yaml(
                    semantic_enabled=semantic_enabled,
                    wiki_enabled=wiki_enabled,
                ),
            )
        )


def test_version_2_needs_no_llm_endpoint_when_semantic_is_disabled(tmp_path: Path) -> None:
    config = load_config(
        _write(tmp_path, _v2_yaml(semantic_enabled=False, wiki_enabled=False))
    )
    assert config.semantic.enabled is False
    assert config.wiki.enabled is False


def test_version_2_rejects_legacy_wiki_settings(tmp_path: Path) -> None:
    value = _v2_yaml(semantic_enabled=False, wiki_enabled=False).replace(
        "wiki:\n  enabled: false",
        "wiki:\n  enabled: false\n  language: Korean",
    )
    with pytest.raises(ValueError, match="unknown wiki setting.*language"):
        load_config(_write(tmp_path, value))


def test_version_2_rejects_unknown_coverage_policy(tmp_path: Path) -> None:
    value = _v2_yaml(semantic_enabled=False, wiki_enabled=False).replace(
        "semantic:\n  enabled: false",
        "semantic:\n  enabled: false\n  coverage_policy: partial",
    )
    with pytest.raises(ValueError, match="semantic.coverage_policy must be one of"):
        load_config(_write(tmp_path, value))


def test_version_2_rejects_unknown_kg_profile(tmp_path: Path) -> None:
    value = _v2_yaml(semantic_enabled=False, wiki_enabled=False).replace(
        "semantic:\n  enabled: false",
        "semantic:\n  enabled: false\n  kg_profile: generic_guessing",
    )
    with pytest.raises(ValueError, match="semantic.kg_profile must be one of"):
        load_config(_write(tmp_path, value))


def test_version_2_rejects_unbounded_relationship_inventory(tmp_path: Path) -> None:
    value = _v2_yaml(semantic_enabled=False, wiki_enabled=False).replace(
        "semantic:\n  enabled: false",
        "semantic:\n  enabled: false\n  max_relationships: 4097",
    )
    with pytest.raises(ValueError, match="semantic.max_relationships must be at most 4096"):
        load_config(_write(tmp_path, value))


@pytest.mark.parametrize(
    "stage,semantic_enabled,wiki_enabled",
    [
        ("parsed", False, False),
        ("semantic", True, False),
        ("wiki", True, True),
    ],
)
def test_allow_existing_still_rejects_nonempty_target_stage_before_parsing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    stage: str,
    semantic_enabled: bool,
    wiki_enabled: bool,
) -> None:
    value = _v2_yaml(
        semantic_enabled=semantic_enabled,
        wiki_enabled=wiki_enabled,
    ).replace(
        "  directory: ./results\n",
        "  directory: ./results\n  allow_existing: true\n",
    )
    if semantic_enabled:
        value = value.replace(
            "llm_api: {}\n",
            "llm_api:\n  base_url: http://127.0.0.1:8001/v1\n  model: local-text\n",
        )
    config = load_config(_write(tmp_path, value))
    source = tmp_path / "source.pptx"
    source.write_bytes(b"preflight-only")
    output = tmp_path / "existing-output"
    target = output / stage
    target.mkdir(parents=True)
    marker = target / "keep.txt"
    marker.write_text("unchanged", encoding="utf-8")

    def fail_if_parsing_starts(*args, **kwargs):
        pytest.fail("run_pipeline was called before stage output preflight")

    monkeypatch.setattr(configured_module, "run_pipeline", fail_if_parsing_starts)

    with pytest.raises(ValueError, match="not empty"):
        configured_module.run_configured(source, config, output_override=output)

    assert marker.read_text(encoding="utf-8") == "unchanged"


def test_config_rejects_duplicate_key(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="duplicate YAML key"):
        load_config(_write(tmp_path, _base_yaml() + "wiki:\n  enabled: false\n"))


@pytest.mark.parametrize(
    "fragment,match",
    [
        ("shared: &x {enabled: false}\n", "anchors are not allowed"),
        ("shared: &x {}\ncopy: *x\n", "anchors are not allowed"),
        ("unknown_section: {}\n", "unknown config setting"),
    ],
)
def test_config_rejects_advanced_or_unknown_yaml(tmp_path: Path, fragment: str, match: str) -> None:
    with pytest.raises(ValueError, match=match):
        load_config(_write(tmp_path, _base_yaml() + fragment))


def test_config_rejects_non_finite_number(tmp_path: Path) -> None:
    value = _base_yaml().replace("backend: powerpoint", "backend: powerpoint\n  source_padding_ratio: .nan")
    with pytest.raises(ValueError, match="finite"):
        load_config(_write(tmp_path, value))


def test_enabled_endpoint_requires_present_environment_secret(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("MISSING_TEST_KEY", raising=False)
    value = _base_yaml().replace(
        "vlm_api: {}",
        "vlm_api:\n  base_url: http://127.0.0.1:8000/v1\n  model: vlm\n  api_key_env: MISSING_TEST_KEY",
    ).replace("enabled: false\n  backend: none", "enabled: true\n  backend: openai_vlm")
    with pytest.raises(ValueError, match="MISSING_TEST_KEY.*not set"):
        load_config(_write(tmp_path, value))


def test_literal_and_environment_keys_are_mutually_exclusive(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "vlm_api: {}",
        "vlm_api:\n  api_key: literal\n  api_key_env: TEST_KEY",
    )
    with pytest.raises(ValueError, match="mutually exclusive"):
        load_config(_write(tmp_path, value))


def test_remote_plain_http_requires_two_explicit_network_permissions(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "vlm_api: {}",
        "vlm_api:\n  base_url: http://192.168.1.20:8000/v1\n  model: vlm",
    ).replace("enabled: false\n  backend: none", "enabled: true\n  backend: openai_vlm")
    with pytest.raises(ValueError, match="allow_remote_endpoints"):
        load_config(_write(tmp_path, value))


def test_windows_reserved_output_stem_is_sanitized(tmp_path: Path) -> None:
    settings = OutputSettings(tmp_path)
    assert settings.path_for(tmp_path / "CON.pptx").name == "_CON"


def test_api_key_is_redacted_from_repr() -> None:
    assert "top-secret" not in repr(APISettings(api_key="top-secret"))


def test_local_openai_endpoint_builds_consistent_v1_chat_url(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "vlm_api: {}",
        "vlm_api:\n  base_url: http://127.0.0.1:8000\n  model: local-vlm",
    ).replace("enabled: false\n  backend: none", "enabled: true\n  backend: openai_vlm")
    config = load_config(_write(tmp_path, value))
    adapter = build_ocr_adapter(config)
    assert isinstance(adapter, OpenAICompatibleVLMAdapter)
    assert adapter.url == "http://127.0.0.1:8000/v1/chat/completions"


def test_local_model_profile_needs_no_vlm_endpoint_and_resolves_bundle_paths(tmp_path: Path) -> None:
    profile = tmp_path / "workers" / "ovisocr2"
    python_path = (
        profile / ".venv" / "Scripts" / "python.exe"
        if os.name == "nt"
        else profile / ".venv" / "bin" / "python"
    )
    python_path.parent.mkdir(parents=True)
    python_path.write_text("", encoding="utf-8")
    worker_path = profile / "worker.py"
    worker_path.write_text("", encoding="utf-8")
    (tmp_path / "models" / "ovisocr2").mkdir(parents=True)
    value = _base_yaml().replace(
        "enabled: false\n  backend: none",
        """enabled: true
  backend: local_model
  local_model:
    model: ovisocr2
    workers_directory: ./workers
    models_directory: ./models
    device: cpu
    dtype: float32
""".rstrip(),
    )

    config = load_config(_write(tmp_path, value))
    adapter = build_ocr_adapter(config)

    assert isinstance(adapter, PersistentOCRWorkerAdapter)
    assert adapter.command[0] == str(python_path)
    assert str(tmp_path / "models" / "ovisocr2") in adapter.command


def test_local_model_rejects_unknown_profile(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "enabled: false\n  backend: none",
        "enabled: true\n  backend: local_model\n  local_model:\n    model: arbitrary_repo",
    )
    with pytest.raises(ValueError, match="ocr.local_model.model must be one of"):
        load_config(_write(tmp_path, value))


@pytest.mark.parametrize("model", ["monkeyocr_v2_b", "ovisocr2"])
def test_local_model_enforces_profile_token_limit(tmp_path: Path, model: str) -> None:
    value = _base_yaml().replace(
        "enabled: false\n  backend: none",
        "enabled: true\n"
        "  backend: local_model\n"
        "  local_model:\n"
        f"    model: {model}\n"
        "    max_new_tokens: 32769",
    )
    with pytest.raises(ValueError, match=rf"at most 32768 for {model}"):
        load_config(_write(tmp_path, value))


def test_paddle_local_model_accepts_larger_token_limit(tmp_path: Path) -> None:
    value = _base_yaml().replace(
        "enabled: false\n  backend: none",
        "enabled: true\n"
        "  backend: local_model\n"
        "  local_model:\n"
        "    model: paddleocr_vl_16\n"
        "    max_new_tokens: 65536",
    )
    assert load_config(_write(tmp_path, value)).ocr.local_model.max_new_tokens == 65_536
