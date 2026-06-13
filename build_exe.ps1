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
& $python -m PyInstaller --noconfirm --clean --onefile --windowed --name Bewerbungsverwaltung run.py

Write-Host "Fertig. EXE liegt unter dist\Bewerbungsverwaltung.exe" -ForegroundColor Green
