from __future__ import annotations

import json
from pathlib import Path
import re

from PIL import Image

from pptx_wiki.ocr import OCRBlock, OCRRequest, OCRResult
from pptx_wiki.pipeline import run_pipeline
from pptx_wiki.synthesis import SynthesisConfig
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
    assert result.wiki is not None and result.wiki.index_path.is_file()
    qa = json.loads(result.qa_path.read_text(encoding="utf-8"))
    assert qa["ocr_failures"] == 0
