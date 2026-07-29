[CmdletBinding()]
param(
    [ValidateSet("cu126", "cpu")]
    [string]$TorchBackend = "cu126",

    [string]$PythonExecutable = "",

    [string]$ModelDir = "",

    # Install the environment only. By default the pinned snapshot is
    # downloaded and verified after dependency installation.
    [switch]$SkipDownload,

    [switch]$Recreate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$WorkerRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent (Split-Path -Parent $WorkerRoot)
$VenvRoot = Join-Path $WorkerRoot ".venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
$Requirements = Join-Path $WorkerRoot "requirements.lock.txt"
$DownloadScript = Join-Path $WorkerRoot "download.py"
$FreezeFile = Join-Path $WorkerRoot "installed.freeze.txt"
if (-not $ModelDir) {
    $ModelDir = Join-Path $RepoRoot "models\monkeyocr_v2_b"
}
$ModelDir = [System.IO.Path]::GetFullPath($ModelDir)

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Program,
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $Program $($Arguments -join ' ')"
    }
}

if ($Recreate -and (Test-Path -LiteralPath $VenvRoot)) {
    $ResolvedWorker = [System.IO.Path]::GetFullPath($WorkerRoot)
    $ResolvedVenv = [System.IO.Path]::GetFullPath($VenvRoot)
    if (-not $ResolvedVenv.StartsWith($ResolvedWorker, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a virtual environment outside the worker directory: $ResolvedVenv"
    }
    Remove-Item -LiteralPath $ResolvedVenv -Recurse -Force
}

if (-not (Test-Path -LiteralPath $VenvPython)) {
    if ($PythonExecutable) {
        Invoke-Checked $PythonExecutable "-m" "venv" $VenvRoot
    }
    elseif (Get-Command py -ErrorAction SilentlyContinue) {
        Invoke-Checked "py" "-3.10" "-m" "venv" $VenvRoot
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        Invoke-Checked "python" "-m" "venv" $VenvRoot
    }
    else {
        throw "Python was not found. Install 64-bit Python 3.10, or pass -PythonExecutable."
    }
}

Invoke-Checked $VenvPython "-c" "import sys; assert sys.maxsize > 2**32, '64-bit Python is required'; raise SystemExit(0 if (3,10) <= sys.version_info[:2] < (3,13) else 'Python 3.10-3.12 is required')"
Invoke-Checked $VenvPython "-m" "pip" "install" "--upgrade" "pip"

if ($TorchBackend -eq "cu126") {
    $TorchIndex = "https://download.pytorch.org/whl/cu126"
}
else {
    $TorchIndex = "https://download.pytorch.org/whl/cpu"
    Write-Warning "CPU inference is intended as a fallback and can be extremely slow for long OCR output."
}

$CurrentTorch = (& $VenvPython -c "import importlib.util; print(__import__('torch').__version__ if importlib.util.find_spec('torch') else '')" 2>$null).Trim()
$ExpectedBuildSuffix = "+$TorchBackend"
$TorchArguments = @(
    "-m", "pip", "install", "--index-url", $TorchIndex,
    "torch==2.6.0", "torchvision==0.21.0"
)
if ($CurrentTorch -and -not $CurrentTorch.Contains($ExpectedBuildSuffix)) {
    # `torch==2.6.0` also accepts local versions such as +cpu and +cu126.
    # Force replacement when the requested wheel backend changed.
    $TorchArguments += "--force-reinstall"
}
Invoke-Checked -Program $VenvPython -Arguments $TorchArguments

Invoke-Checked $VenvPython "-m" "pip" "install" "-r" $Requirements
Invoke-Checked $VenvPython "-m" "pip" "check"

$ExpectCuda = if ($TorchBackend -eq "cpu") { "0" } else { "1" }
$ProbeCode = @"
import torch
assert torch.__version__.split('+', 1)[0] == "2.6.0", torch.__version__
expect_cuda = bool(int("$ExpectCuda"))
if expect_cuda:
    assert torch.cuda.is_available(), "the CUDA wheel is installed but no usable NVIDIA GPU was detected; check the driver or use -TorchBackend cpu"
print(torch.__version__)
print(torch.cuda.get_device_name(0) if torch.cuda.is_available() else "CPU")
"@
Invoke-Checked $VenvPython "-c" $ProbeCode

if (-not $SkipDownload) {
    Write-Host "Downloading or validating the pinned MonkeyOCRv2-B snapshot..."
    Invoke-Checked $VenvPython $DownloadScript "--model-dir" $ModelDir
}

& $VenvPython -m pip freeze --all | Set-Content -LiteralPath $FreezeFile -Encoding utf8
if ($LASTEXITCODE -ne 0) {
    throw "Unable to write the installed package freeze file."
}

Write-Host "MonkeyOCRv2 worker environment is ready."
Write-Host "Python: $VenvPython"
Write-Host "Torch backend: $TorchBackend"
Write-Host "Resolved packages: $FreezeFile"
Write-Host "Model directory: $ModelDir"
if ($SkipDownload) {
    Write-Host "Model download was skipped. Run: & `"$VenvPython`" `"$DownloadScript`" --model-dir `"$ModelDir`""
}
else {
    Write-Host "The pinned model snapshot was downloaded and validated."
}
