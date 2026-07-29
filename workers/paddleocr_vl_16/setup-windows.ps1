[CmdletBinding()]
param(
    [ValidateSet("cpu", "cu118", "cu126", "cu129")]
    [string]$Runtime = "cu126",

    [string]$Python = "py",

    [string]$PythonVersion = "3.11",

    [string]$ModelDir = "",

    [switch]$SkipDownload
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$Lines
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($Path, $Lines, $encoding)
}

if ($env:OS -ne "Windows_NT") {
    throw "This setup script must be run on 64-bit Windows."
}

$WorkerRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent (Split-Path -Parent $WorkerRoot)
if ([string]::IsNullOrWhiteSpace($ModelDir)) {
    $ModelDir = Join-Path $RepoRoot "models\paddleocr_vl_16"
}
$ModelDir = [System.IO.Path]::GetFullPath($ModelDir)
$VenvDir = Join-Path $WorkerRoot ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$RuntimeManifest = Join-Path $WorkerRoot ".runtime.json"
$Requirements = Join-Path $WorkerRoot "requirements.lock.txt"
$FreezePath = Join-Path $WorkerRoot "requirements.freeze.txt"

if (Test-Path $RuntimeManifest) {
    $PreviousRuntime = (Get-Content -Raw -LiteralPath $RuntimeManifest | ConvertFrom-Json).runtime
    if ($PreviousRuntime -ne $Runtime) {
        throw "This worker venv was created for '$PreviousRuntime', not '$Runtime'. Delete '$VenvDir' and '$RuntimeManifest' before changing runtimes."
    }
}

if (-not (Test-Path $VenvPython)) {
    $LauncherArgs = @()
    $PythonLeaf = Split-Path -Leaf $Python
    if ($PythonLeaf -in @("py", "py.exe")) {
        $LauncherArgs += "-$PythonVersion"
    }
    $LauncherArgs += @("-m", "venv", $VenvDir)
    Invoke-Checked -FilePath $Python -Arguments $LauncherArgs
}

Invoke-Checked -FilePath $VenvPython -Arguments @(
    "-c",
    "import platform,sys; assert sys.maxsize > 2**32, '64-bit Python is required'; assert (3,10) <= sys.version_info[:2] <= (3,13), 'Python 3.10-3.13 is required'; print(platform.python_version())"
)

$PaddleVersion = "3.3.1"
if ($Runtime -eq "cpu") {
    Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "uninstall", "-y", "paddlepaddle-gpu")
    $PaddlePackage = "paddlepaddle"
    $PaddleIndex = "https://www.paddlepaddle.org.cn/packages/stable/cpu/"
}
else {
    Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "uninstall", "-y", "paddlepaddle")
    $PaddlePackage = "paddlepaddle-gpu"
    $PaddleIndex = "https://www.paddlepaddle.org.cn/packages/stable/$Runtime/"
}

Invoke-Checked -FilePath $VenvPython -Arguments @(
    "-m", "pip", "install", "$PaddlePackage==$PaddleVersion", "--index-url", $PaddleIndex
)
Invoke-Checked -FilePath $VenvPython -Arguments @(
    "-m", "pip", "install", "--requirement", $Requirements
)
Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "check")

$ExpectCuda = if ($Runtime -eq "cpu") { "0" } else { "1" }
$VerifyCode = @"
import paddle
assert paddle.__version__ == "$PaddleVersion", paddle.__version__
expect_cuda = bool(int("$ExpectCuda"))
assert paddle.device.is_compiled_with_cuda() == expect_cuda, "wrong Paddle runtime installed"
if expect_cuda:
    assert paddle.device.cuda.device_count() > 0, "no usable NVIDIA GPU was detected"
paddle.utils.run_check()
"@
Invoke-Checked -FilePath $VenvPython -Arguments @("-c", $VerifyCode)

if (-not $SkipDownload) {
    Invoke-Checked -FilePath $VenvPython -Arguments @(
        (Join-Path $WorkerRoot "download.py"), "--model-dir", $ModelDir
    )
}

$Frozen = @(& $VenvPython -m pip freeze --all)
if ($LASTEXITCODE -ne 0) {
    throw "pip freeze failed with exit code $LASTEXITCODE"
}
Write-Utf8NoBom -Path $FreezePath -Lines $Frozen

$RuntimeInfo = [ordered]@{
    schema_version = 1
    worker = "paddleocr_vl_16"
    runtime = $Runtime
    paddle_package = $PaddlePackage
    paddle_version = $PaddleVersion
    python = (& $VenvPython -c "import platform; print(platform.python_version())")
    model_dir = $ModelDir
}
$RuntimeJson = $RuntimeInfo | ConvertTo-Json
Write-Utf8NoBom -Path $RuntimeManifest -Lines @($RuntimeJson)

Write-Host "PaddleOCR-VL 1.6 worker is ready."
Write-Host "Python: $VenvPython"
Write-Host "Models: $ModelDir"
Write-Host "Frozen environment: $FreezePath"
