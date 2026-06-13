param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root '.venv\Scripts\python.exe'
if (!(Test-Path $python)) {
    throw "Python in .venv nicht gefunden: $python"
}

if ($Clean) {
    Remove-Item -Recurse -Force "$root\build" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$root\dist" -ErrorAction SilentlyContinue
}

& $python -m pip install -r "$root\requirements.txt"

Write-Host "[Build] Markdown-Lintfix ausführen ..." -ForegroundColor Cyan
& "$root\scripts\markdown_lint_fix.ps1"

$pyiArgs = @("-m", "PyInstaller", "--noconfirm")
if ($Clean) {
    $pyiArgs += "--clean"
}
$pyiArgs += "$root\Bewerbungsverwaltung.spec"
& $python @pyiArgs

Write-Host "Fertig. EXE liegt unter dist\Bewerbungsverwaltung.exe" -ForegroundColor Green
