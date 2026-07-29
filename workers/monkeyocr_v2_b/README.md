# MonkeyOCRv2-B local worker

This worker downloads `zenosai/MonkeyOCRv2-B-Parsing` once and runs it locally
through native Windows PyTorch. It does not call the Hugging Face Inference API
and it does not need vLLM. Runtime networking and Hugging Face telemetry are
forced off.

The worker is intended for already-isolated PowerPoint object/ROI images:

- native PowerPoint text and tables should still be read from OOXML;
- one table ROI is sent with `task: table`, so a neighbouring table cannot be
  merged into it by a second full-slide layout pass;
- the full-slide `document` mode is a fallback, not the default for closely
  spaced PowerPoint objects.

## Requirements

- 64-bit Windows 10/11
- Python 3.10 recommended
- NVIDIA GPU with a current driver for the `cu126` option
- roughly 3 GB of disk space for the complete snapshot and metadata

The checkpoint is about 2.07 GB in total; `model.safetensors` is about 1.76 GB.
The model publisher has not stated a minimum VRAM figure. An 8 GB NVIDIA GPU is
a reasonable starting point for batch size 1, but that is an engineering
estimate, not a vendor guarantee. CPU mode is available as a compatibility
fallback and may be extremely slow.

## 1. Install the independent environment and model

Open PowerShell in this directory:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
./setup-windows.ps1 -TorchBackend cu126
```

For a machine without an NVIDIA GPU:

```powershell
./setup-windows.ps1 -TorchBackend cpu
```

The script creates `workers/monkeyocr_v2_b/.venv`, records the fully resolved
packages in `installed.freeze.txt`, downloads the pinned snapshot into
`<pptx-wiki>/models/monkeyocr_v2_b`, and verifies it. A second run reuses the
downloaded data and verifies it again. It intentionally does not install
`flash-attn`; the pinned model supports PyTorch SDPA on Windows.

To use a non-default model directory:

```powershell
./setup-windows.ps1 -TorchBackend cu126 -ModelDir 'D:\AI\models\monkeyocr_v2_b'
```

To install only the Python environment and postpone the model download:

```powershell
./setup-windows.ps1 -TorchBackend cu126 -SkipDownload
```

If an existing environment was created with the wrong Python version or Torch
backend, add `-Recreate`. This removes and recreates only this worker's `.venv`;
the downloaded model directory is preserved.

## 2. Manually verify or resume the model download (optional)

Normal setup already performs this step. The default target is
`<pptx-wiki>/models/monkeyocr_v2_b`:

```powershell
./.venv/Scripts/python.exe ./download.py
```

Or specify the same custom path used during setup:

```powershell
./.venv/Scripts/python.exe ./download.py `
  --model-dir 'D:\AI\models\monkeyocr_v2_b'
```

The downloader pins this immutable revision:

```text
de7a993bd0f39a97b122dac767e82ae04935bce4
```

It verifies the expected size and SHA-256 of the custom model code, tokenizer,
configuration, and model weights, then writes
`pptx-wiki-model-manifest.json`. To check an existing download without network
access:

```powershell
./.venv/Scripts/python.exe ./download.py --verify-only
```

The model is public, so a Hugging Face token is normally unnecessary. If Hub
access in your environment requires one, put it in `HF_TOKEN`; the downloader
never prints it.

## 3. One-shot smoke test

```powershell
./.venv/Scripts/python.exe ./worker.py `
  --model-dir '..\..\models\monkeyocr_v2_b' `
  --device auto `
  --dtype auto `
  --max-new-tokens 16384 `
  --image 'D:\samples\table-roi.png' `
  --task table `
  --language ko `
  --request-id smoke-1 `
  --output 'D:\samples\table-result.json'
```

`--output` receives a plain common OCR-result JSON object. Protocol/status lines
on stdout always start with `@@PPTX_WIKI@@`; third-party logs go to stderr.

Supported tasks are `document`, `text`, `table`, `chart`, and `formula`.
MonkeyOCR returns tables as OTSL; this worker converts them to escaped HTML and
retains `rowspan`/`colspan` in each table block.

## 4. Persistent JSONL mode

Loading a VLM for every ROI is prohibitively expensive. The deck pipeline
should start one persistent process and reuse it:

```powershell
./.venv/Scripts/python.exe ./worker.py `
  --serve `
  --model-dir '..\..\models\monkeyocr_v2_b' `
  --device auto `
  --dtype auto `
  --max-new-tokens 16384
```

The CLI default is `--max-new-tokens 16384`, matching the root YAML default.
Values up to 32768 are accepted within the checkpoint's 40960-position context;
actual image and prompt tokens also consume that context. Device values accepted
by the root config (`auto`, `cpu`, `cuda`, `cuda:N`, `gpu`, and `gpu:N`) are all
accepted here; `gpu` is normalized to PyTorch's `cuda` spelling.

After model loading and integrity checks, the process emits one ready line:

```text
@@PPTX_WIKI@@{"type":"ready","protocol":"pptx-wiki-ocr-worker/1","ok":true,...}
```

Write one UTF-8 JSON object per line to stdin:

```json
{"protocol":"pptx-wiki-ocr-worker/1","type":"request","op":"recognize","id":"slide-4-table-2","image":"D:\\work\\roi\\table-2.png","image_path":"D:\\work\\roi\\table-2.png","task":"table","language":"ko","context":"2026년 매출 표"}
```

`type` and `op` may be omitted for direct use; they default to a recognition
request. Either `image` or `image_path` is accepted. When both are supplied,
their values must match.

Each response is also one line:

```text
@@PPTX_WIKI@@{"type":"result","protocol":"pptx-wiki-ocr-worker/1","id":"slide-4-table-2","ok":true,"result":{...}}
```

The `result` object contains `text`, `markdown`, `html`, ordered `blocks`,
pixel-relative bboxes, warnings, and raw model output. Request failures return
`ok:false` without terminating the worker. To stop it cleanly:

```json
{"protocol":"pptx-wiki-ocr-worker/1","type":"shutdown","op":"shutdown","id":"shutdown-1"}
```

## Runtime notes

- `--dtype auto` uses BF16 on a compatible NVIDIA GPU, FP16 on an older CUDA
  GPU, and FP32 on CPU. The official checkpoint forces BF16 in its vision path;
  the worker applies a small dtype compatibility wrapper only for FP16/FP32 and
  reports that in `warnings`.
- On CPU the worker uses the checkpoint's eager vision-attention implementation.
  Its SDPA implementation explicitly selects a fused CUDA backend that is not
  available in the PyTorch CPU build.
- Deterministic OCR uses greedy decoding. If the decoder falls into a repeated
  sequence, the worker retries up to three times with gradually raised
  temperature, following the official parser's strategy.
- The geometry/dewarping `.pth` preprocessors are part of the complete snapshot
  but are deliberately not applied to clean, PowerPoint-rendered ROIs. Warping a
  straight slide can damage table alignment.
- This checkpoint requires `trust_remote_code=True`. The worker allows it only
  after checking the pinned local files and then loads with
  `local_files_only=True` while `HF_HUB_OFFLINE=1` and
  `TRANSFORMERS_OFFLINE=1` are set.
- Model weights and the publisher's code are marked Apache-2.0 and the model card
  states that research and commercial use are permitted. Preserve the upstream
  notices when redistributing them.

Official sources:

- [MonkeyOCRv2 repository](https://github.com/Yuliang-Liu/MonkeyOCRv2)
- [MonkeyOCRv2-B-Parsing model card](https://huggingface.co/zenosai/MonkeyOCRv2-B-Parsing)
- [vLLM Windows support note](https://docs.vllm.ai/en/v0.22.0/getting_started/installation/gpu/)
