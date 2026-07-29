from __future__ import annotations

from pathlib import Path

import pytest

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


def _write(tmp_path: Path, value: str) -> Path:
    path = tmp_path / "config.yml"
    path.write_text(value, encoding="utf-8")
    return path


def test_load_config_resolves_paths_from_config_directory(tmp_path: Path) -> None:
    config = load_config(_write(tmp_path, _base_yaml()))
    assert config.output.directory == (tmp_path / "results").resolve()
    assert config.render.backend == "powerpoint"
    assert config.ocr.enabled is False


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
    python_path = profile / ".venv" / "bin" / "python"
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
