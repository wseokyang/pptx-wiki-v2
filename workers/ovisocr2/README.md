# OvisOCR2 offline worker (Windows)

This worker downloads the official `ATH-MaaS/OvisOCR2` checkpoint once and
runs it locally. It does not call a Hugging Face inference endpoint, a paid OCR
API, or the user's OpenAI-compatible API.

The snapshot is pinned to commit
`65c619d374b55d4152e85150fc1b003700bc1f0c` (2026-07-16). It contains a single
1.706 GB BF16 safetensors checkpoint with 852,985,920 parameters and is licensed
under Apache-2.0. The download command checks the checkpoint's size and SHA-256,
validates its safetensors header, and writes a local manifest. Inference refuses
an unmanifested or modified checkpoint and uses `trust_remote_code=False`.

## Why this uses Transformers on Windows

The model card's optimized example pins `vllm==0.22.1`, but vLLM's official
prerequisite is Linux. The native Windows path therefore uses the model's
official Hugging Face Transformers interface with PyTorch SDPA. No optional Hub
kernel, `causal_conv1d`, `fla`, Triton, or remote model code is loaded. Qwen3.5
falls back to built-in PyTorch operations, so it is less optimized than vLLM but
works without WSL. A CUDA-capable NVIDIA GPU is strongly recommended; CPU mode
is intended for functional testing and will be slow.

Primary references:

- [Official OvisOCR2 model card](https://huggingface.co/ATH-MaaS/OvisOCR2)
- [Qwen3.5 Transformers documentation](https://huggingface.co/docs/transformers/model_doc/qwen3_5)
- [Hugging Face snapshot download API](https://huggingface.co/docs/huggingface_hub/main/en/package_reference/file_download)
- [Official PyTorch Windows installer](https://pytorch.org/get-started/locally/)
- [vLLM prerequisites](https://docs.vllm.ai/en/stable/getting_started/quickstart/)

## 1. Install the environment and model

Open PowerShell in this directory. Python 3.12 is required.

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup-windows.ps1 -TorchBackend cu130
```

By default this creates the dedicated environment, installs the selected
PyTorch build, downloads the pinned model into
`<repository>\models\ovisocr2`, and validates every required file. A second run
reuses downloaded data and validates it again.

Available backends are `cpu`, `cu126`, `cu130`, and `cu132`. Choose a CUDA
wheel supported by the installed NVIDIA driver. The script deliberately fails
when a CUDA wheel is selected but PyTorch cannot see a GPU.

To put the model somewhere else:

```powershell
.\setup-windows.ps1 -TorchBackend cu130 -ModelDir "D:\AI\models\ovisocr2"
```

To install only the Python environment and postpone the 1.7 GB model download:

```powershell
.\setup-windows.ps1 -TorchBackend cu130 -SkipDownload
```

The environment is isolated at `.venv`. Installed versions are recorded in
`environment.freeze.txt` and hardware/setup details in `environment.json`.
`-Recreate` removes and recreates only this worker's `.venv`; it does not
delete the model directory.

## 2. Manually validate or re-download the checkpoint (optional)

Normal setup already performs this step. Running `download.py` again resumes
missing files, reuses unchanged files, and repeats the manifest, safetensors,
size, and checksum validation:

```powershell
.\.venv\Scripts\python.exe .\download.py `
  --model-dir "..\..\models\ovisocr2"
```

To fetch the pinned files again without deleting the destination:

```powershell
.\.venv\Scripts\python.exe .\download.py `
  --model-dir "..\..\models\ovisocr2" `
  --force-download
```

The model is public and ungated, so no token is normally required. If a token
is necessary for the local network configuration, keep it out of command
history and put it in an environment variable. This also works before the
normal `setup-windows.ps1` command because setup invokes `download.py`:

```powershell
$env:HF_TOKEN = "hf_..."
.\setup-windows.ps1 -TorchBackend cu130 -ModelDir "D:\AI\models\ovisocr2"
```

For a nonstandard token variable during a manual download:

```powershell
$env:MY_HF_TOKEN = "hf_..."
.\.venv\Scripts\python.exe .\download.py `
  --model-dir "D:\AI\models\ovisocr2" `
  --token-env MY_HF_TOKEN
```

Only the download phase needs network access. The token is neither written to
the manifest nor needed by the inference worker.

## 3. One-shot smoke test

```powershell
.\.venv\Scripts\python.exe .\worker.py `
  --model-dir "..\..\models\ovisocr2" `
  --device cuda `
  --dtype auto `
  --max-new-tokens 16384 `
  --image "D:\samples\table.png" `
  --task table `
  --language ko
```

The first stdout line is `ready`; the second is the OCR `result`. Every
protocol line begins with `@@PPTX_WIKI@@`. Logs, warnings, progress messages,
and tracebacks are written only to stderr.

## Persistent JSONL protocol

The PPTX pipeline should launch one process for the whole deck so the 0.9B
model is loaded only once:

```powershell
.\.venv\Scripts\python.exe .\worker.py `
  --serve `
  --model-dir "..\..\models\ovisocr2" `
  --device auto `
  --dtype auto `
  --max-new-tokens 16384
```

`--device` accepts `auto`, `cpu`, `cuda`, `cuda:N`, `gpu`, and `gpu:N`;
`gpu` spellings from the root YAML are normalized to the corresponding CUDA
device.

Successful startup emits one prefixed line whose JSON payload has this shape:

```json
{"protocol":"pptx-wiki-ocr-worker/1","type":"ready","ok":true,"backend":"hf_ovisocr2"}
```

After `ready`, write one UTF-8 JSON object per stdin line. The bundled main
adapter sends both `image` and the compatibility alias `image_path`; either key
is accepted:

```json
{"protocol":"pptx-wiki-ocr-worker/1","type":"request","op":"recognize","id":"slide-7-table-2","image":"D:\\work\\roi.png","image_path":"D:\\work\\roi.png","task":"table","language":"ko","context":null}
```

The worker writes one corresponding result line and flushes it immediately:

```json
{"protocol":"pptx-wiki-ocr-worker/1","type":"result","id":"slide-7-table-2","ok":true,"result":{"text":"...","markdown":"<table>...</table>","html":"<table>...</table>","confidence":null,"blocks":[]}}
```

`task` must be `document`, `text`, `table`, `chart`, or `formula`. `id` is
returned unchanged. `context` is accepted for pipeline provenance but is not
put into the model prompt. This avoids allowing untrusted slide metadata to
change OCR instructions. Send `{"type":"shutdown"}` or close stdin to stop the
worker. `type=ping` and `op=ping` are also supported for a liveness check.

On startup failure the process emits `type=ready`, `ok=false`, with an error and
then exits nonzero, which lets the main adapter report a configuration error
immediately. A bad request or inference failure emits `type=result`,
`ok=false`, and an `error` object without terminating the persistent process.

## Inference behavior

The worker follows the official model-card settings:

- greedy decoding and thinking disabled;
- `min_pixels = 448 x 448` and `max_pixels = 2880 x 2880`;
- at most 16,384 new tokens by default;
- natural-reading-order Markdown;
- HTML `<table>` output, LaTeX formulas, and visual-region `<img>` tags;
- no translation or paraphrasing.

The visual-region coordinates produced by OvisOCR2 use its normalized
`[0, 1000)` space. The worker converts them to the input ROI's pixel space in
individual blocks. OvisOCR2 does not provide calibrated recognition confidence,
so confidence fields remain `null`.

At runtime the process sets `HF_HUB_OFFLINE=1`, `TRANSFORMERS_OFFLINE=1`, and
related telemetry/progress flags before importing model libraries. It loads
only the validated local directory with `local_files_only=True`.

If CUDA runs out of memory, first lower `--max-pixels` (for example to
`4194304`, or 2048 squared) and then lower `--max-new-tokens`. There is no
official minimum VRAM figure for this checkpoint, so capacity should be tested
with the densest real table slide rather than inferred from weight size alone.
