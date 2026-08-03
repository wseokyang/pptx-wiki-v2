from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path

import pytest

from pptx_wiki.semantic import load_semantic_documents
from pptx_wiki.wiki_publish import WikiExport, publish_wiki


def _json_line(value: dict[str, object]) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True) + "\n"


def _write_artifacts(tmp_path: Path) -> tuple[Path, Path, dict[str, object]]:
    parsed = tmp_path / "parsed" / "corpus"
    slides = parsed / "slides"
    slides.mkdir(parents=True)
    provenance_text = "".join(
        (
            _json_line(
                {
                    "citation": "[slide-1#text-1]",
                    "slide_number": 1,
                    "element_id": "text-1",
                    "content": "매출은 100입니다.",
                }
            ),
            _json_line(
                {
                    "citation": "[slide-2#table-1]",
                    "slide_number": 2,
                    "element_id": "table-1",
                    "content": "| 항목 | 값 |\n| --- | --- |\n| A | 100 |",
                }
            ),
        )
    )
    (parsed / "provenance.jsonl").write_text(
        provenance_text, encoding="utf-8", newline="\n"
    )
    (slides / "slide-0001.md").write_text(
        '# Slide 1\n\n<a id="text-1"></a>\n', encoding="utf-8", newline="\n"
    )
    (slides / "slide-0002.md").write_text(
        '# Slide 2\n\n<a id="table-1"></a>\n', encoding="utf-8", newline="\n"
    )

    semantic = tmp_path / "semantic"
    semantic.mkdir()
    first_body = "- 매출은 100입니다. [slide-1#text-1]"
    second_body = "- 표의 A 값은 100입니다. [slide-2#table-1]"
    documents = [
        {
            "id": "sales",
            "title": "매출 개요",
            "body_markdown": first_body,
            "citations": ["[slide-1#text-1]"],
            "content_sha256": sha256(first_body.encode("utf-8")).hexdigest(),
        },
        {
            "id": "table-summary",
            "title": "표 요약",
            "body_markdown": second_body,
            "citations": ["[slide-2#table-1]"],
            "content_sha256": sha256(second_body.encode("utf-8")).hexdigest(),
        },
    ]
    documents_text = "".join(_json_line(document) for document in documents)
    (semantic / "documents.jsonl").write_text(
        documents_text, encoding="utf-8", newline="\n"
    )
    manifest = {
        "schema_version": "pptx-wiki.semantic.v1",
        "source_provenance_sha256": sha256(
            provenance_text.encode("utf-8")
        ).hexdigest(),
        "documents_file": "documents.jsonl",
        "documents_sha256": sha256(documents_text.encode("utf-8")).hexdigest(),
        "document_count": len(documents),
    }
    (semantic / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return semantic, parsed, manifest


def _rewrite_manifest(semantic: Path, manifest: dict[str, object]) -> None:
    (semantic / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def _rewrite_documents(
    semantic: Path, manifest: dict[str, object], documents: list[dict[str, object]]
) -> None:
    value = "".join(_json_line(document) for document in documents)
    (semantic / "documents.jsonl").write_text(value, encoding="utf-8", newline="\n")
    manifest["documents_sha256"] = sha256(value.encode("utf-8")).hexdigest()
    manifest["document_count"] = len(documents)
    _rewrite_manifest(semantic, manifest)


def test_publish_wiki_is_deterministic_and_links_parsed_slides(tmp_path: Path) -> None:
    semantic, parsed, _ = _write_artifacts(tmp_path)
    empty_output = tmp_path / "wiki-a"
    empty_output.mkdir()

    first = publish_wiki(semantic, parsed, empty_output)
    second = publish_wiki(semantic, parsed, tmp_path / "wiki-b")

    assert isinstance(first, WikiExport)
    assert first.page_count == 2
    assert [path.name for path in first.page_paths] == ["sales.md", "table-summary.md"]
    assert first.index_path.name == "index.md"
    assert first.report_path.name == "publish-report.json"
    page = (empty_output / "sales.md").read_text(encoding="utf-8")
    assert "# 매출 개요" in page
    assert "- 매출은 100입니다. [slide-1#text-1]" in page
    assert (
        "[슬라이드 원문](../parsed/corpus/slides/slide-0001.md#text-1)" in page
    )
    assert "[매출 개요](sales.md)" in first.index_path.read_text(encoding="utf-8")

    for name in ("sales.md", "table-summary.md", "index.md", "publish-report.json"):
        assert (empty_output / name).read_bytes() == (tmp_path / "wiki-b" / name).read_bytes()
    report = json.loads(first.report_path.read_text(encoding="utf-8"))
    assert report["schema_version"] == "pptx-wiki.wiki.v1"
    assert report["page_count"] == 2
    assert [item["file"] for item in report["pages"]] == [
        "sales.md",
        "table-summary.md",
    ]
    assert second.page_count == first.page_count


def test_publish_wiki_rejects_nonempty_or_file_output(tmp_path: Path) -> None:
    semantic, parsed, _ = _write_artifacts(tmp_path)
    nonempty = tmp_path / "nonempty"
    nonempty.mkdir()
    (nonempty / "keep.txt").write_text("keep", encoding="utf-8")

    with pytest.raises(ValueError, match="not empty"):
        publish_wiki(semantic, parsed, nonempty)
    assert (nonempty / "keep.txt").read_text(encoding="utf-8") == "keep"

    output_file = tmp_path / "output-file"
    output_file.write_text("keep", encoding="utf-8")
    with pytest.raises(ValueError, match="not a directory"):
        publish_wiki(semantic, parsed, output_file)
    assert output_file.read_text(encoding="utf-8") == "keep"


def test_publish_wiki_rejects_source_provenance_hash_mismatch(tmp_path: Path) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    manifest["source_provenance_sha256"] = "0" * 64
    _rewrite_manifest(semantic, manifest)
    output = tmp_path / "wiki"

    with pytest.raises(ValueError, match="provenance SHA-256 mismatch"):
        publish_wiki(semantic, parsed, output)
    assert not output.exists()


def test_publish_wiki_rejects_documents_file_hash_mismatch(tmp_path: Path) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    manifest["documents_sha256"] = "f" * 64
    _rewrite_manifest(semantic, manifest)

    with pytest.raises(ValueError, match="documents SHA-256 mismatch"):
        publish_wiki(semantic, parsed, tmp_path / "wiki")


def test_publish_wiki_rejects_document_body_hash_mismatch(tmp_path: Path) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    documents = [
        {
            "id": "tampered",
            "title": "변조됨",
            "body_markdown": "변조된 본문",
            "citations": ["[slide-1#text-1]"],
            "content_sha256": "0" * 64,
        }
    ]
    _rewrite_documents(semantic, manifest, documents)

    with pytest.raises(ValueError, match="content SHA-256 mismatch"):
        publish_wiki(semantic, parsed, tmp_path / "wiki")


def test_publish_wiki_rejects_unknown_citation(tmp_path: Path) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    body = "확인되지 않은 내용"
    documents = [
        {
            "id": "unknown-source",
            "title": "알 수 없는 출처",
            "body_markdown": body,
            "citations": ["[slide-99#missing]"],
            "content_sha256": sha256(body.encode("utf-8")).hexdigest(),
        }
    ]
    _rewrite_documents(semantic, manifest, documents)
    output = tmp_path / "wiki"

    with pytest.raises(ValueError, match="unknown citation"):
        publish_wiki(semantic, parsed, output)
    assert not output.exists()


def test_semantic_loader_and_publisher_reject_undeclared_body_citation(
    tmp_path: Path,
) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    body = "Known source [slide-1#text-1]; fabricated source [slide-99#fabricated]"
    documents = [
        {
            "id": "undeclared-source",
            "title": "Undeclared source",
            "body_markdown": body,
            "citations": ["[slide-1#text-1]"],
            "content_sha256": sha256(body.encode("utf-8")).hexdigest(),
        }
    ]
    _rewrite_documents(semantic, manifest, documents)
    output = tmp_path / "wiki"

    with pytest.raises(ValueError, match="undeclared citation"):
        load_semantic_documents(semantic)
    with pytest.raises(ValueError, match="undeclared citation"):
        publish_wiki(semantic, parsed, output)
    assert not output.exists()


def test_semantic_loader_and_publisher_reject_manifest_document_count_mismatch(
    tmp_path: Path,
) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    manifest["document_count"] = 3
    _rewrite_manifest(semantic, manifest)
    output = tmp_path / "wiki"

    with pytest.raises(ValueError, match="document count"):
        load_semantic_documents(semantic)
    with pytest.raises(ValueError, match="document count"):
        publish_wiki(semantic, parsed, output)
    assert not output.exists()


def test_publish_wiki_rejects_unsafe_or_colliding_document_ids(tmp_path: Path) -> None:
    semantic, parsed, manifest = _write_artifacts(tmp_path)
    body = "본문"
    unsafe = [
        {
            "id": "../escape",
            "title": "탈출",
            "body_markdown": body,
            "citations": ["[slide-1#text-1]"],
            "content_sha256": sha256(body.encode("utf-8")).hexdigest(),
        }
    ]
    _rewrite_documents(semantic, manifest, unsafe)
    with pytest.raises(ValueError, match="unsafe document id"):
        load_semantic_documents(semantic)
    with pytest.raises(ValueError, match="unsafe document id"):
        publish_wiki(semantic, parsed, tmp_path / "wiki")

    duplicate = [
        {
            "id": value,
            "title": value,
            "body_markdown": body,
            "citations": ["[slide-1#text-1]"],
            "content_sha256": sha256(body.encode("utf-8")).hexdigest(),
        }
        for value in ("Topic", "topic")
    ]
    _rewrite_documents(semantic, manifest, duplicate)
    with pytest.raises(ValueError, match="duplicate semantic document id"):
        load_semantic_documents(semantic)
    with pytest.raises(ValueError, match="duplicate semantic document id"):
        publish_wiki(semantic, parsed, tmp_path / "wiki")
