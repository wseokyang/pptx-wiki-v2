from __future__ import annotations

import os
from pathlib import Path
import subprocess

import pytest

from pptx_wiki.quartz_publish import publish_quartz


def _make_directory_link(link: Path, target: Path) -> None:
    """Create a directory symlink, or an unprivileged junction on Windows."""

    target = target.resolve(strict=True)
    try:
        link.symlink_to(target, target_is_directory=True)
        return
    except OSError as symlink_error:
        if os.name != "nt":
            pytest.skip(f"directory symlink is unavailable: {symlink_error}")

    completed = subprocess.run(
        ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)],
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.returncode != 0:
        pytest.skip(
            "directory symlink/junction is unavailable: "
            + (completed.stderr.strip() or completed.stdout.strip())
        )
    assert link.is_dir()


def test_publish_quartz_rejects_collection_with_linked_ancestor(
    tmp_path: Path,
) -> None:
    real_parent = tmp_path / "real-input-parent"
    collection = real_parent / "collection"
    integrated = collection / "integrated"
    integrated.mkdir(parents=True)

    linked_parent = tmp_path / "linked-input-parent"
    _make_directory_link(linked_parent, real_parent)

    with pytest.raises(ValueError, match=r"(?i)(?:symlink|junction|reparse)"):
        publish_quartz(
            linked_parent / "collection",
            linked_parent / "collection" / "integrated",
            tmp_path / "quartz-output",
        )


def test_publish_quartz_rejects_linked_output_directory(tmp_path: Path) -> None:
    collection = tmp_path / "collection"
    integrated = collection / "integrated"
    integrated.mkdir(parents=True)

    real_output = tmp_path / "real-quartz-output"
    real_output.mkdir()
    linked_output = tmp_path / "linked-quartz-output"
    _make_directory_link(linked_output, real_output)

    with pytest.raises(ValueError, match=r"(?i)(?:symlink|junction|reparse)"):
        publish_quartz(
            collection,
            integrated,
            linked_output,
        )
