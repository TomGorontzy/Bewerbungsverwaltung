param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseVersion,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

if ($ReleaseVersion -notmatch '^v\d+\.\d+\.\d+$') {
    throw "Ungültiges Versionsformat '$ReleaseVersion'. Erwartet z. B. v0.1.1"
}

Write-Host "[Release] Markdown-Lintfix ausführen ..." -ForegroundColor Cyan
& "$root\scripts\markdown_lint_fix.ps1"

$statusArgs = @('status', '--porcelain')
$mdChanges = @((& git @statusArgs) | Where-Object { $_ -imatch '\.(md|markdown|mdown|mkd)$' })
if ($mdChanges.Count -gt 0) {
    Write-Host "[Release] Markdown-Änderungen gefunden, committe automatisch ..." -ForegroundColor Yellow
    $mdPaths = @()
    foreach ($line in $mdChanges) {
        $pathPart = $line.Substring(3).Trim()
        $renameSeparator = ' -' + '> '
        if ($pathPart.Contains($renameSeparator)) {
            $pathPart = $pathPart.Split($renameSeparator)[1]
        }
        $mdPaths += $pathPart
    }

    foreach ($path in ($mdPaths | Sort-Object -Unique)) {
        $addArgs = @('add', "$path")
        & git @addArgs
    }
    $commitArgs = @('commit', '-m', "Docs: Markdown-Lintfix vor Release $ReleaseVersion")
    & git @commitArgs
    $pushArgs = @('push')
    & git @pushArgs
}

Write-Host "[Release] EXE bauen ..." -ForegroundColor Cyan
if ($Clean) {
    & "$root\build_exe.ps1" -Clean
} else {
    & "$root\build_exe.ps1"
}

Write-Host "[Release] Tag erstellen und pushen ..." -ForegroundColor Cyan
$tagArgs = @('tag', '-a', $ReleaseVersion, '-m', "Release $ReleaseVersion")
& git @tagArgs
$pushTagArgs = @('push', 'origin', $ReleaseVersion)
& git @pushTagArgs

Write-Host "[Release] GitHub Release erstellen ..." -ForegroundColor Cyan
$releaseArgs = @('release', 'create', $ReleaseVersion, '--title', $ReleaseVersion, '--generate-notes')

# Assets sammeln
$assets = @(
    'dist/Bewerbungsverwaltung.exe',
    'Bewerbungsaktivitäten mit Erinnerungen.xlsx',
    'docs/DOKUMENTATION_ANWENDER.md',
    'docs/DOKUMENTATION_TECHNIK.md',
    'docs/FAQ.md',
    'README.md'
)

foreach ($asset in $assets) {
    if (Test-Path $asset) {
        $releaseArgs += $asset
    } else {
        Write-Host "[Release] ⚠ Asset nicht gefunden: $asset" -ForegroundColor Yellow
    }
}

& gh @releaseArgs

Write-Host "[Release] Fertig: $ReleaseVersion" -ForegroundColor Green
Write-Host "[Release] Release-URL: https://github.com/TomGorontzy/Bewerbungsverwaltung/releases/tag/$ReleaseVersion" -ForegroundColor Cyan
