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

Write-Host "[Markdown] Starte Auto-Fix ..." -ForegroundColor Cyan
& $python -m pymarkdown -d $disabledRules fix $mdFiles
if ($LASTEXITCODE -ne 0) {
    throw "Markdown-Fix fehlgeschlagen (ExitCode $LASTEXITCODE)."
}

Write-Host "[Markdown] Starte Verifikation ..." -ForegroundColor Cyan
& $python -m pymarkdown -d $disabledRules scan $mdFiles
if ($LASTEXITCODE -ne 0) {
    throw "Markdown-Scan fehlgeschlagen (ExitCode $LASTEXITCODE)."
}

Write-Host "[Markdown] Lint/Fix abgeschlossen." -ForegroundColor Green
