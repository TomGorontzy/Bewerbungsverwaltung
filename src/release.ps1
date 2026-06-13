param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseVersion,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$srcRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $srcRoot
Set-Location $projectRoot

$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

if ($ReleaseVersion -notmatch '^v\d+\.\d+\.\d+$') {
    throw "Ungültiges Versionsformat '$ReleaseVersion'. Erwartet z. B. v0.1.1"
}

Write-Host "[Release] Markdown-Lintfix ausführen ..." -ForegroundColor Cyan
& "$srcRoot\scripts\markdown_lint_fix.ps1"

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
    & "$srcRoot\build_exe.ps1" -Clean
} else {
    & "$srcRoot\build_exe.ps1"
}

Write-Host "[Release] ZIP-Archiv erstellen ..." -ForegroundColor Cyan

# Releases-Verzeichnis vorbereiten
$releasesDir = Join-Path $projectRoot 'releases'
if (!(Test-Path $releasesDir)) {
    New-Item $releasesDir -ItemType Directory -Force | Out-Null
    Write-Host "[Release] Releases-Verzeichnis erstellt: $releasesDir" -ForegroundColor DarkGreen
}

# ZIP vorbereiten
$zipName = "Bewerbungsverwaltung-${ReleaseVersion}.zip"
$zipPath = Join-Path $releasesDir $zipName
$tempDir = Join-Path $projectRoot "release-temp"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item $tempDir -ItemType Directory | Out-Null

# Dateien ins temp-Verzeichnis kopieren
@(
    'dist/Bewerbungsverwaltung.exe',
    'README.md'
) | ForEach-Object {
    if (Test-Path $_) {
        Copy-Item $_ $tempDir -Force
        Write-Host "  + $(Split-Path -Leaf $_)" -ForegroundColor DarkGreen
    }
}

# Excel-Datei unter data/ bereitstellen
$dataSource = Join-Path $projectRoot 'data\Bewerbungsaktivitäten mit Erinnerungen.xlsx'
$dataDestDir = Join-Path $tempDir 'data'
if (Test-Path $dataSource) {
    New-Item $dataDestDir -ItemType Directory -Force | Out-Null
    Copy-Item $dataSource $dataDestDir -Force
    Write-Host "  + data/Bewerbungsaktivitäten mit Erinnerungen.xlsx" -ForegroundColor DarkGreen
} else {
    throw "Excel-Datei nicht gefunden: $dataSource"
}

# Docs-Verzeichnis mit allen Markdown-Dateien
$docsSource = Join-Path $projectRoot 'docs'
$docsDest = Join-Path $tempDir 'docs'
if (Test-Path $docsSource) {
    New-Item $docsDest -ItemType Directory -Force | Out-Null
    Get-ChildItem -Path $docsSource -Filter '*.md' -File | ForEach-Object {
        Copy-Item $_.FullName $docsDest -Force
        Write-Host "  + docs/$($_.Name)" -ForegroundColor DarkGreen
    }
}

# ZIP erstellen
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Host "[Release] ZIP erstellt: $zipName" -ForegroundColor Green
Write-Host "[Release] Lokal gespeichert: releases/$zipName" -ForegroundColor DarkCyan

# Aufräumen
Remove-Item $tempDir -Recurse -Force

# WICHTIG: Commits JETZT sammeln, VOR der Tag-Erstellung
# (Nach Tag-Erstellung würde git describe das neue Tag zurückgeben)
Write-Host "[Release] Commits für Tag-Nachricht und Release-Notes vorbereiten ..." -ForegroundColor Cyan

$previousTag = (& git describe --tags --abbrev=0 2>$null)
if ([string]::IsNullOrWhiteSpace($previousTag)) {
    $previousTag = "v0.0.0"
}
$allCommits = @(git log "$previousTag..HEAD" --oneline 2>$null)
$commitCount = $allCommits.Count

Write-Host "[Release] Tag erstellen und pushen ..." -ForegroundColor Cyan

# Tag-Nachricht vorbereiten
$tagDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$tagMessage = @"
Release: $ReleaseVersion
Datum: $tagDate
Commits: $commitCount

Änderungen seit $previousTag :
"@

if ($commitCount -gt 0) {
    $tagMessage += "`n"
    foreach ($commit in $allCommits) {
        $tagMessage += "  • $commit`n"
    }
} else {
    $tagMessage += "`n  (Keine Commits)"
}

$tagMessage += "`nZIP: $zipName"

# Tag mit ausführlicher Nachricht erstellen
$tagArgs = @('tag', '-a', $ReleaseVersion, '-m', $tagMessage)
& git @tagArgs
$pushTagArgs = @('push', 'origin', $ReleaseVersion)
& git @pushTagArgs

Write-Host "[Release] Tag mit Versionsinformationen erstellt" -ForegroundColor Green

Write-Host "[Release] GitHub Release erstellen ..." -ForegroundColor Cyan

# Release-Notes mit Zusammenfassung vorbereiten
$notesFileName = "RELEASE_NOTES_$ReleaseVersion.md"
$notesFile = Join-Path $releasesDir $notesFileName
$notesContent = @"
# RELEASE NOTES $ReleaseVersion

## Kurzüberblick

Dieses Release enthält die seit dem vorherigen Tag übernommenen Änderungen.

## Enthaltene Änderungen

"@

# Commits analysieren und kategorisieren
$commitTypes = @{
    'fix:'       = '🔧 Fehlerbehebungen'
    'feature:'   = '✨ Neue Features'
    'release:'   = '📦 Release & Build'
    'docs:'      = '📚 Dokumentation'
    'chore:'     = '🔨 Maintenance'
    'refactor:'  = '♻️ Refactoring'
    'ci:'        = '🚀 CI/CD'
}

# Bereits gesammelte Commits kategorisieren (nicht erneut abrufen!)
$categorized = @{}
foreach ($commit in $allCommits) {
    # Format: <hash> <type>: <message>
    $parts = $commit -split '\s+', 2
    if ($parts.Count -eq 2) {
        $message = $parts[1]

        $found = $false
        foreach ($typeKey in $commitTypes.Keys) {
            if ($message -imatch "^$([regex]::Escape($typeKey))") {
                if (-not $categorized.ContainsKey($typeKey)) {
                    $categorized[$typeKey] = @()
                }
                $categorized[$typeKey] += $message
                $found = $true
                break
            }
        }
        if (-not $found) {
            if (-not $categorized.ContainsKey('other')) {
                $categorized['other'] = @()
            }
            $categorized['other'] += $message
        }
    }
}

# Abschnitt "Enthaltene Änderungen" füllen
$hasCategorizedEntries = $false
foreach ($typeKey in @('fix:', 'feature:', 'release:', 'docs:', 'chore:', 'refactor:', 'ci:', 'other')) {
    if ($categorized.ContainsKey($typeKey) -and $categorized[$typeKey].Count -gt 0) {
        $label = if ($typeKey -eq 'other') { 'Sonstiges' } else { $commitTypes[$typeKey] }
        $notesContent += "`n### $label`n"
        foreach ($msg in $categorized[$typeKey]) {
            $notesContent += "- $msg`n"
        }
        $hasCategorizedEntries = $true
    }
}

if (-not $hasCategorizedEntries) {
    $notesContent += "`n- Keine kategorisierbaren Änderungen gefunden.`n"
}

$notesContent += "`n## Commit-Übersicht`n`n"
if ($allCommits.Count -gt 0) {
    foreach ($commit in $allCommits) {
        $parts = $commit -split '\s+', 2
        if ($parts.Count -eq 2) {
            $notesContent += "- ``$($parts[0])`` $($parts[1])`n"
        } else {
            $notesContent += "- $commit`n"
        }
    }
} else {
    $notesContent += "- Keine Commits im Vergleichszeitraum gefunden.`n"
}

$notesContent += @"

## Breaking Changes

- Keine bekannten Breaking Changes.

## Migrationshinweise

- Keine Migration erforderlich.

## Known Issues

- Keine neuen bekannten Probleme in diesem Release dokumentiert.

## Hinweise

- Diese Datei wurde automatisch durch `src/release.ps1` erzeugt.

---

Weitere Details unter: [Changelog](https://github.com/TomGorontzy/Bewerbungsverwaltung/commits/$ReleaseVersion)
"@

# Notes-Datei schreiben
[System.IO.File]::WriteAllText($notesFile, $notesContent, [System.Text.UTF8Encoding]::new($false))
Write-Host "[Release] Release-Notes mit Zusammenfassung erstellt: releases/$notesFileName" -ForegroundColor Green

# GitHub Release mit Notes-Datei erstellen
$releaseArgs = @('release', 'create', $ReleaseVersion, $zipPath, '--title', $ReleaseVersion, '--notes-file', $notesFile)
& gh @releaseArgs

Write-Host "[Release] Fertig: $ReleaseVersion" -ForegroundColor Green
Write-Host "[Release] Lokal gespeichert: releases/$zipName" -ForegroundColor DarkCyan
Write-Host "[Release] Release-Notes gespeichert: releases/$notesFileName" -ForegroundColor DarkCyan
Write-Host "[Release] GitHub Release: https://github.com/TomGorontzy/Bewerbungsverwaltung/releases/tag/$ReleaseVersion" -ForegroundColor Cyan
