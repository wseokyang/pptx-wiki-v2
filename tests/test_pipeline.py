from __future__ import annotations

import json
from pathlib import Path
import re

from PIL import Image

import pptx_wiki.configured as configured_module
from pptx_wiki.config import load_config
from pptx_wiki.configured import run_configured
from pptx_wiki.ocr import OCRBlock, OCRRequest, OCRResult
from pptx_wiki.pipeline import run_pipeline
from pptx_wiki.synthesis import SynthesisConfig, WikiSynthesis
from pptx_wiki.wiki_output import load_provenance


class _FixtureOCR:
    name = "fixture"

    def recognize(self, request: OCRRequest) -> OCRResult:
        assert Path(request.image).is_file()
        return OCRResult(
            backend=self.name,
            text="IMAGE-TABLE IMG-A 1,234.50",
            markdown="| 코드 | 값 |\n| --- | --- |\n| IMG-A | 1,234.50 |",
            blocks=[
                OCRBlock(
                    kind="table",
                    text="IMAGE-TABLE IMG-A 1,234.50",
                    markdown="| 코드 | 값 |\n| --- | --- |\n| IMG-A | 1,234.50 |",
                    bbox=(24, 24, 824, 324),
                    confidence=0.9,
                )
            ],
        )


class _GroundedWikiBackend:
    def complete(self, messages, *, max_tokens: int, temperature: float) -> str:
        citations = re.findall(r"\[slide-\d+#[^\]\s#]+\]", messages[-1]["content"])
        assert citations
        return f"# 추출 내용\n\n- 원본 블록을 그대로 근거로 사용합니다. {citations[0]}"


def test_pipeline_native_plus_roi_ocr_without_paid_services(complex_pptx, tmp_path: Path) -> None:
    source, _ = complex_pptx
    slides_dir = tmp_path / "rendered"
    slides_dir.mkdir()
    for number in (1, 2):
        Image.new("RGB", (4000, 2250), "white").save(slides_dir / f"slide-{number:04d}.png")

    result = run_pipeline(
        source,
        tmp_path / "out",
        ocr_adapter=_FixtureOCR(),
        rendered_slides_dir=slides_dir,
        synthesis_backend=_GroundedWikiBackend(),
        synthesis_config=SynthesisConfig(discover_topics=False),
    )

    records = load_provenance(result.corpus.provenance_path)
    assert len([record for record in records if record["kind"] == "table"]) == 2
    assert len([record for record in records if record["kind"] == "ocr_table"]) == 1
    assert result.ocr_successes == 1
    assert result.semantic is not None and result.semantic.manifest_path.is_file()
    assert result.wiki is not None and result.wiki.index_path.is_file()
    assert result.wiki.topic_count == len(result.wiki.topic_paths)
    assert result.wiki.semantic is result.semantic
    qa = json.loads(result.qa_path.read_text(encoding="utf-8"))
    assert qa["ocr_failures"] == 0


def test_configured_pipeline_returns_the_same_wiki_result_contract(
    complex_pptx,
    tmp_path: Path,
    monkeypatch,
) -> None:
    source, _ = complex_pptx
    config_path = tmp_path / "config.yml"
    config_path.write_text(
        """version: 2
output:
  directory: ./unused
render:
  backend: auto
extraction: {}
vlm_api: {}
llm_api:
  base_url: http://127.0.0.1:8000/v1
  model: fixture-llm
  api_key_env: ""
ocr:
  enabled: false
  backend: none
semantic:
  enabled: true
  goal: Keep source facts.
  coverage_policy: selected
  discover_topics: false
wiki:
  enabled: true
network: {}
""",
        encoding="utf-8",
    )
    config = load_config(config_path)
    backend = _GroundedWikiBackend()
    monkeypatch.setattr(
        configured_module,
        "OpenAICompatibleClient",
        lambda **kwargs: backend,
    )

    result = run_configured(source, config, output_override=tmp_path / "configured")

    assert isinstance(result.wiki, WikiSynthesis)
    assert result.semantic is not None
    assert result.wiki.semantic is result.semantic
    assert result.wiki.topic_count == len(result.wiki.topic_paths)
    assert result.wiki.index_path.is_file()
