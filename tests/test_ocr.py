from __future__ import annotations

import json
import sys
import subprocess
from pathlib import Path
from typing import Any, Mapping

import pytest
from PIL import Image

from pptx_wiki.ocr import (
    CommandOCRAdapter,
    OCRConfigurationError,
    OCRExecutionError,
    OCRRequest,
    OCRTransientError,
    OpenAICompatibleVLMAdapter,
    PaddleOCRCLIAdapter,
    PersistentOCRWorkerAdapter,
    RetryPolicy,
    create_ocr_adapter,
    normalize_ocr_payload,
)


PNG_HEADER = b"\x89PNG\r\n\x1a\nnot-decoded-by-fake-backend"


def test_normalize_paddle_keeps_adjacent_tables_separate() -> None:
    payload = {
        "res": {
            "parsing_res_list": [
                {
                    "block_label": "table",
                    "block_content": "<table><tr><td>왼쪽</td></tr></table>",
                    "block_bbox": [0, 0, 90, 100],
                    "block_order": 1,
                },
                {
                    "block_label": "table",
                    "block_content": "<table><tr><td>오른쪽</td></tr></table>",
                    "block_bbox": [94, 0, 190, 100],
                    "block_order": 2,
                },
            ],
            "layout_det_res": {
                "boxes": [
                    {"label": "table", "score": 0.93, "coordinate": [0, 0, 90, 100]},
                    {"label": "table", "score": 0.88, "coordinate": [94, 0, 190, 100]},
                ]
            },
        }
    }

    result = normalize_ocr_payload(payload, backend="paddle_cli")

    assert len(result.blocks) == 2
    assert [block.text for block in result.blocks] == ["왼쪽", "오른쪽"]
    assert [block.confidence for block in result.blocks] == pytest.approx([0.93, 0.88])
    assert result.html is not None and result.html.count("<table>") == 2


def test_normalize_page_markdown_keeps_multiple_html_tables_as_blocks() -> None:
    markdown = (
        "왼쪽 표\n\n<table><tr><td>A</td></tr></table>\n\n"
        "오른쪽 표\n\n<table><tr><td>B</td></tr></table>"
    )

    result = normalize_ocr_payload(
        {"markdown": markdown, "blocks": []}, backend="ovisocr2"
    )

    assert [block.kind for block in result.blocks] == ["text", "table", "text", "table"]
    assert [block.text for block in result.blocks if block.kind == "table"] == ["A", "B"]


def test_paddle_table_roi_disables_layout_detection(tmp_path: Path) -> None:
    seen_command: list[str] = []

    def fake_runner(command: list[str], **_: Any) -> subprocess.CompletedProcess[str]:
        seen_command.extend(command)
        output_dir = Path(command[command.index("--save_path") + 1])
        payload = {
            "res": {
                "parsing_res_list": [
                    {
                        "block_label": "table",
                        "block_content": "<table><tr><td>1</td></tr></table>",
                        "block_bbox": [0, 0, 100, 50],
                        "block_order": 1,
                    }
                ]
            }
        }
        (output_dir / "roi_res.json").write_text(json.dumps(payload), encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, "", "")

    adapter = PaddleOCRCLIAdapter(runner=fake_runner, retry=RetryPolicy(attempts=1))
    result = adapter.recognize(OCRRequest(image=PNG_HEADER, task="table"))

    assert "--use_layout_detection" in seen_command
    assert seen_command[seen_command.index("--use_layout_detection") + 1] == "False"
    assert seen_command[seen_command.index("--prompt_label") + 1] == "table"
    assert result.html == "<table><tr><td>1</td></tr></table>"


def test_document_paddle_command_preserves_distinct_layout_boxes() -> None:
    adapter = PaddleOCRCLIAdapter()
    command = adapter._command(Path("slide.png"), Path("out"), OCRRequest(image=PNG_HEADER))
    assert command[command.index("--merge_layout_blocks") + 1] == "False"
    assert command[command.index("--layout_merge_bboxes_mode") + 1] == "union"


def test_command_adapter_reads_common_json() -> None:
    seen_command: list[str] = []

    def fake_runner(command: list[str], **_: Any) -> subprocess.CompletedProcess[str]:
        seen_command.extend(command)
        output = Path(command[command.index("--output") + 1])
        output.write_text(
            json.dumps(
                {
                    "text": "매출 10%",
                    "markdown": "매출 10%",
                    "html": None,
                    "confidence": 0.8,
                    "blocks": [],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(command, 0, "", "")

    adapter = CommandOCRAdapter(
        ["python", "hf_ocr.py", "--image", "{image}", "--output", "{output}", "--task", "{task}"],
        backend_name="monkeyocr",
        runner=fake_runner,
        retry=RetryPolicy(attempts=1),
    )
    result = adapter.recognize(OCRRequest(image=PNG_HEADER, task="text"))

    assert seen_command[seen_command.index("--task") + 1] == "text"
    assert result.backend == "monkeyocr"
    assert result.text == "매출 10%"
    assert result.confidence == pytest.approx(0.8)


def test_adapter_applies_pixel_crop_before_backend(tmp_path: Path) -> None:
    image_path = tmp_path / "slide.png"
    Image.new("RGB", (20, 10), "white").save(image_path)
    observed_size: list[tuple[int, int]] = []

    def fake_runner(command: list[str], **_: Any) -> subprocess.CompletedProcess[str]:
        roi_path = Path(command[command.index("--image") + 1])
        with Image.open(roi_path) as roi:
            observed_size.append(roi.size)
        output = Path(command[command.index("--output") + 1])
        output.write_text('{"text":"x","markdown":"x","html":null,"blocks":[]}', encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, "", "")

    adapter = CommandOCRAdapter(
        ["ocr", "--image", "{image}", "--output", "{output}"],
        runner=fake_runner,
        retry=RetryPolicy(attempts=1),
    )
    adapter.recognize(OCRRequest(image=image_path, crop=(3, 2, 11, 9), task="text"))

    assert observed_size == [(8, 7)]


def test_command_adapter_does_not_forward_scrubbed_api_key(monkeypatch) -> None:
    observed: list[str | None] = []

    def fake_runner(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        observed.append(kwargs["env"].get("PPTX_WIKI_SECRET"))
        output = Path(command[command.index("--output") + 1])
        output.write_text('{"text":"x","markdown":"x","html":null,"blocks":[]}', encoding="utf-8")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setenv("PPTX_WIKI_SECRET", "secret")
    adapter = CommandOCRAdapter(
        ["ocr", "--image", "{image}", "--output", "{output}"],
        scrub_env_vars=("PPTX_WIKI_SECRET",),
        runner=fake_runner,
        retry=RetryPolicy(attempts=1),
    )
    adapter.recognize(OCRRequest(image=PNG_HEADER, task="text"))

    assert observed == [None]


def test_openai_vlm_retries_and_parses_fenced_json() -> None:
    calls: list[Mapping[str, Any]] = []

    def transport(
        _: str, payload: Mapping[str, Any], __: Mapping[str, str], ___: float
    ) -> Mapping[str, Any]:
        calls.append(payload)
        if len(calls) == 1:
            raise OCRTransientError("model is warming up")
        content = """```json
        {"text":"A","markdown":"| A |","html":"<table><tr><td>A</td></tr></table>",
         "confidence":null,"blocks":[]}
        ```"""
        return {"choices": [{"message": {"content": content}}]}

    adapter = OpenAICompatibleVLMAdapter(
        base_url="http://127.0.0.1:8000/v1",
        model="local-vlm",
        transport=transport,
        retry=RetryPolicy(attempts=2, initial_delay=0),
        response_format="json_schema",
    )
    result = adapter.recognize(OCRRequest(image=PNG_HEADER, task="table", language_hint="ko"))

    assert len(calls) == 2
    assert result.html == "<table><tr><td>A</td></tr></table>"
    user_text = calls[-1]["messages"][1]["content"][0]["text"]
    assert "exactly ONE table" in user_text
    assert calls[-1]["response_format"]["type"] == "json_schema"


def test_factory_has_stable_names() -> None:
    assert isinstance(create_ocr_adapter("paddle_cli"), PaddleOCRCLIAdapter)
    assert isinstance(
        create_ocr_adapter("command", command=["runner", "{image}"]), CommandOCRAdapter
    )
    assert isinstance(
        create_ocr_adapter("openai_vlm", base_url="http://localhost:8000/v1", model="vlm"),
        OpenAICompatibleVLMAdapter,
    )


def test_persistent_worker_loads_once_and_reuses_process(tmp_path: Path, monkeypatch) -> None:
    worker = tmp_path / "fake_worker.py"
    worker.write_text(
        """import json, os, sys
PREFIX = '@@PPTX_WIKI@@'
PROTOCOL = 'pptx-wiki-ocr-worker/1'
print('library startup noise', flush=True)
print(PREFIX + json.dumps({'protocol': PROTOCOL, 'type': 'ready', 'revision': 'test-sha'}), flush=True)
count = 0
for line in sys.stdin:
    request = json.loads(line)
    if request.get('type') == 'shutdown':
        break
    count += 1
    secret = os.getenv('TEST_WORKER_SECRET', '') + os.getenv('OPENAI_API_KEY', '')
    result = {'text': f\"{count}:{request['task']}:{secret}\", 'markdown': 'ok', 'blocks': []}
    envelope = {'protocol': PROTOCOL, 'type': 'result', 'id': request['id'], 'ok': True, 'result': result}
    print(PREFIX + json.dumps(envelope), flush=True)
""",
        encoding="utf-8",
    )
    monkeypatch.setenv("TEST_WORKER_SECRET", "must-not-leak")
    monkeypatch.setenv("OPENAI_API_KEY", "also-must-not-leak")
    adapter = PersistentOCRWorkerAdapter(
        [sys.executable, "-u", str(worker)],
        backend_name="fake_local",
        startup_timeout=5,
        request_timeout=5,
        retry=RetryPolicy(attempts=1),
        scrub_env_vars=("TEST_WORKER_SECRET",),
    )
    try:
        first = adapter.recognize(OCRRequest(image=PNG_HEADER, task="text"))
        second = adapter.recognize(OCRRequest(image=PNG_HEADER, task="table"))
    finally:
        adapter.close()

    assert first.text == "1:text:"
    assert second.text == "2:table:"
    assert adapter.ready_metadata is None


def test_persistent_worker_reports_startup_fatal_without_waiting_for_timeout(
    tmp_path: Path,
) -> None:
    worker = tmp_path / "fatal_worker.py"
    worker.write_text(
        """import json
print('@@PPTX_WIKI@@' + json.dumps({
    'protocol': 'pptx-wiki-ocr-worker/1',
    'type': 'fatal',
    'ok': False,
    'error': {'code': 'startup_failed', 'message': 'bad snapshot'},
}), flush=True)
""",
        encoding="utf-8",
    )
    adapter = PersistentOCRWorkerAdapter(
        [sys.executable, "-u", str(worker)],
        backend_name="fatal_local",
        startup_timeout=5,
        request_timeout=5,
        retry=RetryPolicy(attempts=1),
    )

    with pytest.raises(OCRConfigurationError, match="bad snapshot"):
        adapter.recognize(OCRRequest(image=PNG_HEADER, task="text"))


def test_persistent_worker_restarts_after_retryable_result(tmp_path: Path) -> None:
    attempts = tmp_path / "attempts.txt"
    worker = tmp_path / "retry_worker.py"
    worker.write_text(
        """import json, pathlib, sys
PREFIX = '@@PPTX_WIKI@@'
PROTOCOL = 'pptx-wiki-ocr-worker/1'
counter = pathlib.Path(sys.argv[1])
attempt = int(counter.read_text() or '0') + 1 if counter.exists() else 1
counter.write_text(str(attempt))
print(PREFIX + json.dumps({'protocol': PROTOCOL, 'type': 'ready'}), flush=True)
for line in sys.stdin:
    request = json.loads(line)
    if request.get('type') == 'shutdown':
        break
    if attempt == 1:
        envelope = {
            'protocol': PROTOCOL, 'type': 'result', 'id': request['id'], 'ok': False,
            'error': {'code': 'temporary_failure', 'message': 'retry me', 'retryable': True},
        }
    else:
        envelope = {
            'protocol': PROTOCOL, 'type': 'result', 'id': request['id'], 'ok': True,
            'result': {'text': 'recovered', 'markdown': 'recovered', 'blocks': []},
        }
    print(PREFIX + json.dumps(envelope), flush=True)
""",
        encoding="utf-8",
    )
    adapter = PersistentOCRWorkerAdapter(
        [sys.executable, "-u", str(worker), str(attempts)],
        backend_name="retry_local",
        startup_timeout=5,
        request_timeout=5,
        retry=RetryPolicy(attempts=2, initial_delay=0),
    )
    try:
        result = adapter.recognize(OCRRequest(image=PNG_HEADER, task="text"))
    finally:
        adapter.close()

    assert result.text == "recovered"
    assert attempts.read_text(encoding="utf-8") == "2"


def test_persistent_worker_rejects_mismatched_result_id(tmp_path: Path) -> None:
    worker = tmp_path / "wrong_id_worker.py"
    worker.write_text(
        """import json, sys
prefix = '@@PPTX_WIKI@@'
protocol = 'pptx-wiki-ocr-worker/1'
print(prefix + json.dumps({'protocol': protocol, 'type': 'ready'}), flush=True)
for line in sys.stdin:
    request = json.loads(line)
    if request.get('type') == 'shutdown':
        break
    print(prefix + json.dumps({
        'protocol': protocol, 'type': 'result', 'id': 'another-request',
        'ok': True, 'result': {'text': 'wrong', 'blocks': []},
    }), flush=True)
""",
        encoding="utf-8",
    )
    adapter = PersistentOCRWorkerAdapter(
        [sys.executable, "-u", str(worker)],
        backend_name="wrong_id_local",
        startup_timeout=5,
        request_timeout=5,
        retry=RetryPolicy(attempts=1),
    )

    with pytest.raises(OCRExecutionError, match="unexpected request id"):
        adapter.recognize(
            OCRRequest(image=PNG_HEADER, task="text", request_id="expected-request")
        )
