param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$PptxPath
)

$ErrorActionPreference = "Stop"
$ProjectDir = $PSScriptRoot
$ConfigPath = Join-Path $ProjectDir "config.yml"
$PythonPath = Join-Path $ProjectDir ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $PythonPath -PathType Leaf)) {
    throw "Python environment not found. Run setup-windows.ps1 first: $PythonPath"
}
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Trusted configuration file not found: $ConfigPath"
}
if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) {
    throw "PPTX file not found: $PptxPath"
}
if ([IO.Path]::GetExtension($PptxPath).ToLowerInvariant() -ne ".pptx") {
    throw "Only macro-free .pptx files are accepted: $PptxPath"
}
$ResolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$env:PYTHONUTF8 = "1"

& $PythonPath -m pptx_wiki.cli convert $ResolvedPptx --config $ConfigPath
if ($LASTEXITCODE -ne 0) {
    throw "pptx-wiki failed with exit code $LASTEXITCODE"
}
