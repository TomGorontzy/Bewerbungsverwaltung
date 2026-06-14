param(
    [string[]]$Files = @(
        'DOKUMENTATION_ANWENDER.md',
        'DOKUMENTATION_TECHNIK.md',
        'SCHNELLSTART.md'
    )
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcRoot = Split-Path -Parent $scriptRoot
$projectRoot = Split-Path -Parent $srcRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'

if (!(Test-Path $python)) {
    throw "Python in .venv nicht gefunden: $python"
}

$scriptPath = Join-Path $scriptRoot 'convert_docs_to_pdf.py'
if (!(Test-Path $scriptPath)) {
    throw "Konvertierungsskript nicht gefunden: $scriptPath"
}

$arguments = @($scriptPath, '--project-root', $projectRoot, '--files') + $Files
& $python @arguments

Write-Host "[PDF] Konvertierung abgeschlossen." -ForegroundColor Green
