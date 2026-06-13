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

Write-Host "[Release] ZIP-Archiv erstellen ..." -ForegroundColor Cyan

# Releases-Verzeichnis vorbereiten
$releasesDir = Join-Path $root 'releases'
if (!(Test-Path $releasesDir)) {
    New-Item $releasesDir -ItemType Directory -Force | Out-Null
    Write-Host "[Release] Releases-Verzeichnis erstellt: $releasesDir" -ForegroundColor DarkGreen
}

# ZIP vorbereiten
$zipName = "Bewerbungsverwaltung-${ReleaseVersion}.zip"
$zipPath = Join-Path $releasesDir $zipName
$tempDir = Join-Path $root "release-temp"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item $tempDir -ItemType Directory | Out-Null

# Dateien ins temp-Verzeichnis kopieren
@(
    'dist/Bewerbungsverwaltung.exe',
    'Bewerbungsaktivitäten mit Erinnerungen.xlsx',
    'README.md'
) | ForEach-Object {
    if (Test-Path $_) {
        Copy-Item $_ $tempDir -Force
        Write-Host "  + $(Split-Path -Leaf $_)" -ForegroundColor DarkGreen
    }
}

# Docs-Verzeichnis mit allen Markdown-Dateien
$docsSource = Join-Path $root 'docs'
$docsDest = Join-Path $tempDir 'docs'
if (Test-Path $docsSource) {
    New-Item $docsDest -ItemType Directory -Force | Out-Null
    @('DOKUMENTATION_ANWENDER.md', 'DOKUMENTATION_TECHNIK.md', 'FAQ.md') | ForEach-Object {
        $docFile = Join-Path $docsSource $_
        if (Test-Path $docFile) {
            Copy-Item $docFile $docsDest -Force
            Write-Host "  + docs/$_" -ForegroundColor DarkGreen
        }
    }
}

# ZIP erstellen
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Host "[Release] ZIP erstellt: $zipName" -ForegroundColor Green
Write-Host "[Release] Lokal gespeichert: releases/$zipName" -ForegroundColor DarkCyan

# Aufräumen
Remove-Item $tempDir -Recurse -Force

Write-Host "[Release] Tag erstellen und pushen ..." -ForegroundColor Cyan

# Tag-Nachricht vorbereiten
$tagDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$previousTag = $(git describe --tags --abbrev=0 2>$null) || "v0.0.0"

# Commits seit letztem Tag sammeln
$commitsSinceTag = @(git log "$previousTag..HEAD" --oneline 2>$null)
$commitCount = $commitsSinceTag.Count

$tagMessage = @"
Release: $ReleaseVersion
Datum: $tagDate
Commits: $commitCount

Änderungen seit $previousTag :
"@

if ($commitCount -gt 0) {
    $tagMessage += "`n"
    foreach ($commit in $commitsSinceTag) {
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
$notesFile = Join-Path $root "release-notes-$ReleaseVersion.tmp"
$notesContent = @"
## Zusammenfassung (DE)

Dieses Release beinhaltet:

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

# Commits seit letztem Tag sammeln
$previousTag = $(git describe --tags --abbrev=0 2>$null) || "v0.0.0"
$allCommits = @(git log "$previousTag..HEAD" --oneline 2>$null)

# Commits nach Typ kategorisieren
$categorized = @{}
foreach ($commit in $allCommits) {
    # Format: <hash> <type>: <message>
    $parts = $commit -split '\s+', 2
    if ($parts.Count -eq 2) {
        $hash = $parts[0]
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

# Zusammenfassung füllen
foreach ($typeKey in @('fix:', 'feature:', 'release:', 'docs:', 'chore:', 'refactor:', 'ci:', 'other')) {
    if ($categorized.ContainsKey($typeKey) -and $categorized[$typeKey].Count -gt 0) {
        $label = if ($typeKey -eq 'other') { 'Sonstiges' } else { $commitTypes[$typeKey] }
        $notesContent += "`n- $label`n"
        foreach ($msg in $categorized[$typeKey]) {
            $notesContent += "  - $msg`n"
        }
    }
}

$notesContent += @"

## Hinweise

- Dieser Release fokussiert auf Prozess- und Tooling-Qualität.
- Stabilisierung der Release-/Build-Pipeline mit automatischer Markdown-Lint-Bereinigung.
- Verbesserungen für größere Teams und reproduzierbare Abläufe.

---

Weitere Details unter: [Changelog](https://github.com/TomGorontzy/Bewerbungsverwaltung/commits/$ReleaseVersion)
"@

# Notes-Datei schreiben
[System.IO.File]::WriteAllText($notesFile, $notesContent, [System.Text.UTF8Encoding]::new($false))
Write-Host "[Release] Release-Notes mit Zusammenfassung erstellt" -ForegroundColor Green

# GitHub Release mit Notes-Datei erstellen
$releaseArgs = @('release', 'create', $ReleaseVersion, $zipPath, '--title', $ReleaseVersion, '--notes-file', $notesFile)
& gh @releaseArgs

# Aufräumen
Remove-Item $notesFile -Force

Write-Host "[Release] Fertig: $ReleaseVersion" -ForegroundColor Green
Write-Host "[Release] Lokal gespeichert: releases/$zipName" -ForegroundColor DarkCyan
Write-Host "[Release] GitHub Release: https://github.com/TomGorontzy/Bewerbungsverwaltung/releases/tag/$ReleaseVersion" -ForegroundColor Cyan
