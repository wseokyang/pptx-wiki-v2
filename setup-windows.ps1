$ErrorActionPreference = "Stop"
$ProjectDir = $PSScriptRoot
Set-Location -LiteralPath $ProjectDir
$ConfigPath = Join-Path $ProjectDir "config.yml"
$ExampleConfigPath = Join-Path $ProjectDir "config.example.yml"

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

$Launcher = Get-Command py.exe -ErrorAction SilentlyContinue
if ($null -ne $Launcher) {
    Invoke-Checked -FilePath $Launcher.Source -Arguments @("-3", "-m", "venv", ".venv")
} else {
    $Python = Get-Command python.exe -ErrorAction Stop
    Invoke-Checked -FilePath $Python.Source -Arguments @("-m", "venv", ".venv")
}

$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip")
Invoke-Checked -FilePath $VenvPython -Arguments @("-m", "pip", "install", "-e", ".[api,windows]")

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    if (-not (Test-Path -LiteralPath $ExampleConfigPath -PathType Leaf)) {
        throw "Configuration template not found: $ExampleConfigPath"
    }
    Copy-Item -LiteralPath $ExampleConfigPath -Destination $ConfigPath
    Write-Host "Created local configuration: $ConfigPath"
}

Write-Host ""
Write-Host "Main application setup complete."
Write-Host "1. Install one isolated OCR worker (recommended default):"
Write-Host "   .\workers\paddleocr_vl_16\setup-windows.ps1 -Runtime cu126"
Write-Host "   Use -Runtime cpu when no supported NVIDIA GPU is available."
Write-Host "2. Edit: $ProjectDir\config.yml"
Write-Host '3. If using api_key_env for Wiki synthesis, set it before running, for example:'
Write-Host '   $env:PPTX_WIKI_API_KEY = "your-key"'
Write-Host "4. Run: .\run-windows.ps1 'C:\path\to\deck.pptx'"
