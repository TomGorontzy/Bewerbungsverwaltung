param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$srcRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $srcRoot
Set-Location $projectRoot

$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (!(Test-Path $python)) {
    throw "Python in .venv nicht gefunden: $python"
}

if ($Clean) {
    Remove-Item -Recurse -Force "$projectRoot\build" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$projectRoot\dist" -ErrorAction SilentlyContinue
}

& $python -m pip install -r "$srcRoot\requirements.txt"

Write-Host "[Build] Markdown-Lintfix ausführen ..." -ForegroundColor Cyan
& "$srcRoot\scripts\markdown_lint_fix.ps1"

$pyiArgs = @("-m", "PyInstaller", "--noconfirm")
if ($Clean) {
    $pyiArgs += "--clean"
}
$pyiArgs += "$srcRoot\Bewerbungsverwaltung.spec"
& $python @pyiArgs

Write-Host "Fertig. EXE liegt unter dist\Bewerbungsverwaltung.exe" -ForegroundColor Green
