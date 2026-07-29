[CmdletBinding()]
param(
    [ValidateSet("cpu", "cu126", "cu130", "cu132")]
    [string]$TorchBackend = "cu130",

    # The default uses the Python launcher and creates a Python 3.12 venv.
    # Supply an absolute python.exe path when the launcher is unavailable.
    [string]$PythonCommand = "py",

    # Empty means <repository>\models\ovisocr2.
    [string]$ModelDir = "",

    # Install the environment only. By default the pinned model is downloaded
    # and validated after dependency installation.
    [switch]$SkipDownload,

    [switch]$Recreate
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$WorkerDir = [System.IO.Path]::GetFullPath($PSScriptRoot)
$RepoRoot = [System.IO.Path]::GetFullPath((Join-Path $WorkerDir "..\.."))
if ([string]::IsNullOrWhiteSpace($ModelDir)) {
    $ModelDir = Join-Path $RepoRoot "models\ovisocr2"
}
$ModelDir = [System.IO.Path]::GetFullPath($ModelDir)
$VenvDir = Join-Path $WorkerDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$Requirements = Join-Path $WorkerDir "requirements.lock.txt"
$DownloadScript = Join-Path $WorkerDir "download.py"
$FreezePath = Join-Path $WorkerDir "environment.freeze.txt"
$SetupRecordPath = Join-Path $WorkerDir "environment.json"

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

if ($env:OS -ne "Windows_NT") {
    throw "This setup script must be run on 64-bit Windows."
}

if (-not (Test-Path -LiteralPath $Requirements -PathType Leaf)) {
    throw "Missing dependency lock file: $Requirements"
}
if (-not (Test-Path -LiteralPath $DownloadScript -PathType Leaf)) {
    throw "Missing model downloader: $DownloadScript"
}

if ($Recreate -and (Test-Path -LiteralPath $VenvDir)) {
    # VenvDir is deliberately fixed below this worker directory.  Refuse any
    # unexpected path before recursively removing it.
    $ExpectedVenv = [System.IO.Path]::GetFullPath((Join-Path $WorkerDir ".venv"))
    if ([System.IO.Path]::GetFullPath($VenvDir) -ne $ExpectedVenv) {
        throw "Refusing to remove an unexpected virtual environment path: $VenvDir"
    }
    Remove-Item -LiteralPath $VenvDir -Recurse -Force
}

if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
    $PythonLeaf = Split-Path -Leaf $PythonCommand
    if ($PythonLeaf -in @("py", "py.exe")) {
        Invoke-Checked -FilePath $PythonCommand -Arguments @("-3.12", "-m", "venv", $VenvDir)
    }
    else {
        Invoke-Checked -FilePath $PythonCommand -Arguments @("-m", "venv", $VenvDir)
    }
}

$PythonVersion = (& $VenvPython -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')").Trim()
if ($LASTEXITCODE -ne 0) {
    throw "Unable to run the worker virtual environment: $VenvPython"
}
if ($PythonVersion -ne "3.12") {
    throw "OvisOCR2 worker requires Python 3.12; the venv contains Python $PythonVersion. Re-run with -Recreate."
}

$TorchIndexes = @{
    "cpu"   = "https://download.pytorch.org/whl/cpu"
    "cu126" = "https://download.pytorch.org/whl/cu126"
    "cu130" = "https://download.pytorch.org/whl/cu130"
    "cu132" = "https://download.pytorch.org/whl/cu132"
}
$TorchIndex = $TorchIndexes[$TorchBackend]

Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip")
$CurrentTorch = (& $VenvPython -c "import importlib.util; print(__import__('torch').__version__ if importlib.util.find_spec('torch') else '')" 2>$null).Trim()
$ExpectedBuildSuffix = "+$TorchBackend"
$TorchArguments = @(
    "-m", "pip", "install", "--index-url", $TorchIndex,
    "torch==2.13.0", "torchvision==0.28.0"
)
if ($CurrentTorch -and -not $CurrentTorch.Contains($ExpectedBuildSuffix)) {
    # PEP 440 considers 2.13.0+cpu and 2.13.0+cu130 compatible with
    # `torch==2.13.0`, so force replacement when switching wheel backends.
    $TorchArguments += "--force-reinstall"
}
Invoke-Checked -FilePath $VenvPython -Arguments $TorchArguments
Invoke-Checked -FilePath $VenvPython -Arguments @(
    "-m", "pip", "install", "--requirement", $Requirements
)
Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "check")

$ProbeCode = @'
import json
import platform
import sys
import torch
import torchvision
import transformers
import huggingface_hub

value = {
    "python": platform.python_version(),
    "implementation": platform.python_implementation(),
    "architecture": platform.machine(),
    "torch": torch.__version__,
    "torchvision": torchvision.__version__,
    "transformers": transformers.__version__,
    "huggingface_hub": huggingface_hub.__version__,
    "cuda_available": torch.cuda.is_available(),
    "torch_cuda": torch.version.cuda,
    "cuda_device": torch.cuda.get_device_name(0) if torch.cuda.is_available() else None,
}
print(json.dumps(value, ensure_ascii=False))
'@
$ProbeJson = (& $VenvPython -c $ProbeCode).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "The installed OvisOCR2 environment failed its import check."
}
$Probe = $ProbeJson | ConvertFrom-Json

if ($TorchBackend -ne "cpu" -and -not $Probe.cuda_available) {
    throw "Installed the $TorchBackend build, but PyTorch cannot access an NVIDIA GPU. Check the NVIDIA driver or use -TorchBackend cpu."
}

$FreezeLines = & $VenvPython -m pip freeze --all
if ($LASTEXITCODE -ne 0) {
    throw "Unable to record the installed environment."
}
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines($FreezePath, [string[]]$FreezeLines, $Utf8NoBom)

$SetupRecord = [ordered]@{
    created_at_utc = [DateTime]::UtcNow.ToString("o")
    torch_backend = $TorchBackend
    torch_index = $TorchIndex
    worker_directory = $WorkerDir
    virtual_environment = $VenvDir
    model_directory = $ModelDir
    model_download_skipped = [bool]$SkipDownload
    packages = $Probe
}

if (-not $SkipDownload) {
    Write-Host "Downloading or validating the pinned OvisOCR2 model..."
    Invoke-Checked -FilePath $VenvPython -Arguments @($DownloadScript, "--model-dir", $ModelDir)
}

[System.IO.File]::WriteAllText(
    $SetupRecordPath,
    ($SetupRecord | ConvertTo-Json -Depth 5),
    $Utf8NoBom
)

Write-Host "OvisOCR2 worker environment is ready."
Write-Host "Python: $VenvPython"
Write-Host "Backend: $TorchBackend"
Write-Host "Environment record: $SetupRecordPath"
Write-Host "Model directory: $ModelDir"
if ($SkipDownload) {
    Write-Host "Model download was skipped. Run: & `"$VenvPython`" `"$DownloadScript`" --model-dir `"$ModelDir`""
}
else {
    Write-Host "The pinned model was downloaded and validated."
}
