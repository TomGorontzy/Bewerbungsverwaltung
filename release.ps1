param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseVersion,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if ($ReleaseVersion -notmatch '^v\d+\.\d+\.\d+$') {
    throw "Ungültiges Versionsformat '$ReleaseVersion'. Erwartet z. B. v0.1.1"
}

Write-Host "[Release] Markdown-Lintfix ausführen ..." -ForegroundColor Cyan
& "$root\scripts\markdown_lint_fix.ps1"

$mdChanges = git status --porcelain -- '*.md'
if ($mdChanges) {
    Write-Host "[Release] Markdown-Änderungen gefunden, committe automatisch ..." -ForegroundColor Yellow
    git add '*.md'
    git commit -m "Docs: Markdown-Lintfix vor Release $ReleaseVersion"
    git push
}

Write-Host "[Release] EXE bauen ..." -ForegroundColor Cyan
if ($Clean) {
    & "$root\build_exe.ps1" -Clean
} else {
    & "$root\build_exe.ps1"
}

Write-Host "[Release] Tag erstellen und pushen ..." -ForegroundColor Cyan
git tag -a $ReleaseVersion -m "Release $ReleaseVersion"
git push origin $ReleaseVersion

Write-Host "[Release] GitHub Release erstellen ..." -ForegroundColor Cyan
gh release create $ReleaseVersion "dist/Bewerbungsverwaltung.exe" --title "$ReleaseVersion" --generate-notes

Write-Host "[Release] Fertig: $ReleaseVersion" -ForegroundColor Green
