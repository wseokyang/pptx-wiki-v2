import json
from pathlib import Path
from types import SimpleNamespace
from typing import Mapping, Sequence

import pptx_wiki.cli as cli_module
from pptx_wiki.cli import main
from pptx_wiki.models import BBox, DeckRecord, Element, SlideRecord
from pptx_wiki.semantic import SemanticConfig, build_semantic_output
from pptx_wiki.wiki_output import export_slide_corpus


class _GroundedBackend:
    model = "test-model"

    def __init__(self) -> None:
        self.calls = 0

    def complete(
        self,
        messages: Sequence[Mapping[str, str]],
        *,
        max_tokens: int,
        temperature: float,
    ) -> str:
        self.calls += 1
        return "# Quarterly Results\n\n- Revenue is 100. [slide-1#revenue]"


def _write_parsed_fixture(root: Path) -> Path:
    parsed = root / "parsed"
    slide = SlideRecord(
        number=1,
        width=1_000,
        height=1_000,
        title="Quarterly Results",
        elements=[
            Element(
                id="revenue",
                slide_number=1,
                kind="text",
                bbox=BBox(10, 10, 400, 100),
                z_index=0,
                text="Revenue is 100.",
            )
        ],
    )
    export_slide_corpus(
        DeckRecord(
            source_path="quarterly-results.pptx",
            slide_width=slide.width,
            slide_height=slide.height,
            slides=[slide],
        ),
        parsed / "corpus",
    )
    return parsed


def test_cli_coverage_defaults_keep_compatibility_and_favor_selection() -> None:
    parser = cli_module._parser()

    legacy = parser.parse_args(["run", "source.pptx", "-o", "out", "--synthesize"])
    organize = parser.parse_args(["organize", "parsed", "-o", "semantic"])

    assert legacy.coverage_policy == "complete"
    assert organize.coverage_policy == "selected"


def test_cli_parse_creates_only_the_parsed_stage(
    complex_pptx, tmp_path: Path, capsys
) -> None:
    source, _ = complex_pptx
    output = tmp_path / "parse-output"

    assert main(["parse", str(source), "-o", str(output)]) == 0

    assert (output / "parsed" / "manifest.json").is_file()
    assert (output / "parsed" / "deck.json").is_file()
    assert (output / "parsed" / "qa.json").is_file()
    assert (output / "parsed" / "corpus" / "provenance.jsonl").is_file()
    assert not (output / "semantic").exists()
    assert not (output / "wiki").exists()
    summary = json.loads(capsys.readouterr().out)
    assert summary["slides"] == 2
    assert summary["semantic_documents"] == 0
    assert summary["wiki_pages"] == 0


def test_cli_organize_creates_semantic_artifact(
    tmp_path: Path, capsys, monkeypatch
) -> None:
    parsed = _write_parsed_fixture(tmp_path / "workspace")
    semantic = tmp_path / "semantic"
    backend = _GroundedBackend()
    monkeypatch.setattr(cli_module, "_semantic_backend", lambda args: backend)

    assert (
        main(
            [
                "organize",
                str(parsed),
                "-o",
                str(semantic),
                "--no-topic-discovery",
                "--repair-attempts",
                "0",
            ]
        )
        == 0
    )

    assert (semantic / "manifest.json").is_file()
    assert (semantic / "documents.jsonl").is_file()
    summary = json.loads(capsys.readouterr().out)
    assert summary["documents"] == 1
    assert summary["selected_blocks"] == 1
    assert summary["omitted_blocks"] == 0
    assert backend.calls == 1


def test_cli_wiki_publishes_semantic_artifact_from_sibling_parsed_dir(
    tmp_path: Path, capsys
) -> None:
    workspace = tmp_path / "workspace"
    parsed = _write_parsed_fixture(workspace)
    semantic = workspace / "semantic"
    wiki = workspace / "wiki"
    build_semantic_output(
        parsed / "corpus",
        backend=_GroundedBackend(),
        output_dir=semantic,
        config=SemanticConfig(discover_topics=False, repair_attempts=0),
    )

    assert main(["wiki", str(semantic), "-o", str(wiki)]) == 0

    assert (wiki / "index.md").is_file()
    assert (wiki / "quarterly-results.md").is_file()
    assert (wiki / "publish-report.json").is_file()
    summary = json.loads(capsys.readouterr().out)
    assert summary["pages"] == 1
    assert Path(summary["wiki_dir"]) == wiki.resolve()


def test_cli_quartz_dispatches_without_loading_config_or_collection_pipeline(
    tmp_path: Path, capsys, monkeypatch
) -> None:
    collection = tmp_path / "finished-collection"
    collection.mkdir()
    output = tmp_path / "republished-quartz"
    recorded: dict[str, object] = {}

    def forbidden(*args, **kwargs):
        raise AssertionError("Quartz resume must not load config or rerun collection")

    def fake_publish_quartz(
        collection_dir: Path,
        integrated_dir: Path,
        output_dir: Path,
        *,
        site_title: str,
    ) -> SimpleNamespace:
        recorded.update(
            collection_dir=collection_dir,
            integrated_dir=integrated_dir,
            output_dir=output_dir,
            site_title=site_title,
        )
        return SimpleNamespace(
            output_dir=output_dir,
            content_dir=output_dir / "content",
            page_count=7,
            pr_count=2,
            entity_count=3,
            relationship_count=4,
            asset_paths=(),
        )

    monkeypatch.setattr(cli_module, "load_config", forbidden)
    monkeypatch.setattr(cli_module, "run_configured_collection", forbidden)
    monkeypatch.setattr(cli_module, "publish_quartz", fake_publish_quartz)

    assert (
        main(
            [
                "quartz",
                str(collection),
                "--output",
                str(output),
                "--site-title",
                "Recovered Wiki",
            ]
        )
        == 0
    )

    assert recorded == {
        "collection_dir": collection.resolve(),
        "integrated_dir": collection.resolve() / "integrated",
        "output_dir": output.resolve(),
        "site_title": "Recovered Wiki",
    }
    summary = json.loads(capsys.readouterr().out)
    assert summary == {
        "quartz_dir": str(output.resolve()),
        "content_dir": str(output.resolve() / "content"),
        "pages": 7,
        "prs": 2,
        "entities": 3,
        "relationships": 4,
        "assets": 0,
    }


def test_cli_runs_native_only(complex_pptx, tmp_path: Path, capsys) -> None:
    source, _ = complex_pptx
    output = tmp_path / "cli-output"

    assert main(["run", str(source), "-o", str(output)]) == 0

    assert (output / "parsed" / "manifest.json").is_file()
    assert (output / "parsed" / "deck.json").is_file()
    assert (output / "parsed" / "qa.json").is_file()
    assert (output / "parsed" / "corpus" / "provenance.jsonl").is_file()
    assert '"slides": 2' in capsys.readouterr().out


def test_cli_convert_reads_trusted_yaml_and_needs_only_pptx_input(
    complex_pptx, tmp_path: Path, capsys
) -> None:
    source, _ = complex_pptx
    config_path = tmp_path / "config.yml"
    config_path.write_text(
        """version: 1
output:
  directory: ./configured-output
  subdirectory_per_pptx: true
  naming: stem
  allow_existing: false
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
""",
        encoding="utf-8",
    )

    assert main(["convert", str(source), "--config", str(config_path)]) == 0

    output = tmp_path / "configured-output" / source.stem
    assert (output / "parsed" / "deck.json").is_file()
    console = capsys.readouterr().out
    assert "Preflight (secrets redacted)" in console
    assert "Completed" in console
