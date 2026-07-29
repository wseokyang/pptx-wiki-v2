from pathlib import Path

from pptx_wiki.cli import main


def test_cli_runs_native_only(complex_pptx, tmp_path: Path, capsys) -> None:
    source, _ = complex_pptx
    output = tmp_path / "cli-output"

    assert main(["run", str(source), "-o", str(output)]) == 0

    assert (output / "deck.json").is_file()
    assert (output / "qa.json").is_file()
    assert (output / "corpus" / "provenance.jsonl").is_file()
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
    assert (output / "deck.json").is_file()
    console = capsys.readouterr().out
    assert "Preflight (secrets redacted)" in console
    assert "Completed" in console
