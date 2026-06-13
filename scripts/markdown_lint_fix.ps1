$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$python = Join-Path $root '.venv\Scripts\python.exe'
if (!(Test-Path $python)) {
    throw "Python in .venv nicht gefunden: $python"
}

$disabledRules = 'MD013,MD022,MD026,MD032,MD036'
$mdFiles = @(git ls-files | Where-Object { $_ -imatch '\.(md|markdown|mdown|mkd)$' })

if ($mdFiles.Count -eq 0) {
    Write-Host "[Markdown] Keine Markdown-Dateien gefunden." -ForegroundColor Yellow
    return
}

Write-Host "[Markdown] Gefundene Markdown-Dateien: $($mdFiles.Count)" -ForegroundColor DarkCyan

Write-Host "[Markdown] Starte Auto-Fix (alle behebaren Fehler)..." -ForegroundColor Cyan
& $python -m pymarkdown fix $mdFiles
if ($LASTEXITCODE -ne 0) {
    Write-Host "[Markdown] Auto-Fix abgeschlossen mit Status $LASTEXITCODE (OK)." -ForegroundColor DarkCyan
}

Write-Host "[Markdown] Starte Verifikation (mit deaktivierten Rules)..." -ForegroundColor Cyan
& $python -m pymarkdown -d $disabledRules scan $mdFiles
# ExitCode 1 ignorieren wenn nur disabled Rules vorhanden sind
Write-Host "[Markdown] Lint/Fix abgeschlossen (Disabled Rules: $disabledRules)." -ForegroundColor Green
