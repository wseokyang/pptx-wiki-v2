"""Pinned MonkeyOCRv2-B-Parsing model metadata.

The checkpoint contains custom Python code and therefore must never be loaded
from a moving branch.  These values pin both the source revision and every file
that is required before ``trust_remote_code=True`` is allowed.
"""

from __future__ import annotations

MODEL_REPO_ID = "zenosai/MonkeyOCRv2-B-Parsing"
MODEL_REVISION = "de7a993bd0f39a97b122dac767e82ae04935bce4"
MANIFEST_FILENAME = "pptx-wiki-model-manifest.json"

# path: (size in bytes, sha256)
REQUIRED_FILES: dict[str, tuple[int, str]] = {
    "added_tokens.json": (
        707,
        "c0284b582e14987fbd3d5a2cb2bd139084371ed9acbae488829a1c900833c680",
    ),
    "chat_template.jinja": (
        5_292,
        "3636d0f0bd6bef02654cdffdc447b79cb2cef8ab02cc75267345946291a489e4",
    ),
    "config.json": (
        2_186,
        "2b756ea5479f3fda8ca058132cd8fb8851a489a7c77538082fb7ff0540d8ebd2",
    ),
    "configuration_monkeyocrv2.py": (
        3_156,
        "c045b7476f1d2278a9953438ca89a0383849cf46ffe286989680b314d006be1e",
    ),
    "generation_config.json": (
        214,
        "d3e057bbca66b92f33a8bdc6a1301014e0e4ab69b3b3fd2e442d9fe0c69f3431",
    ),
    "merges.txt": (
        1_671_853,
        "8831e4f1a044471340f7c0a83d7bd71306a5b867e95fd870f74d0c5308a904d5",
    ),
    "model.safetensors": (
        1_755_925_032,
        "0267fdc991c9be02cf1b60405f77fe0629d084970ad9d6d08163feda3d470284",
    ),
    "modeling_monkeyocrv2.py": (
        5_550,
        "3a9e1f2508c9ccac71f12b080fbaf135264c625ff4191cb06696b1a43793b960",
    ),
    "modeling_monkeyocrv2_vision.py": (
        19_941,
        "411161a04945e36f60217b72e39b1ccdc2c309a1190c2750f01679a51c0eb3aa",
    ),
    "preprocessor_config.json": (
        449,
        "1beeee7ca6fb814296224d57de413990c5488c277aea43bf0ca16f110fa38a64",
    ),
    "processor_config.json": (
        139,
        "246cfb3914c6c277e29686a02583e9895d6b993d0c2c51adc175b0819bd70d57",
    ),
    "special_tokens_map.json": (
        613,
        "76862e765266b85aa9459767e33cbaf13970f327a0e88d1c65846c2ddd3a1ecd",
    ),
    "tokenizer.json": (
        11_422_654,
        "aeb13307a71acd8fe81861d94ad54ab689df773318809eed3cbe794b4492dae4",
    ),
    "tokenizer_config.json": (
        5_540,
        "aa854f394b95a30114bb1f45632175426d75683b0fc27437f4daebeeb81dd020",
    ),
    "vocab.json": (
        2_776_833,
        "ca10d7e9fb3ed18575dd1e277a2579c16d108e32f27439684afa0e10b1440910",
    ),
}
