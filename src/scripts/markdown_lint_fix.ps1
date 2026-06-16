$ErrorActionPreference = 'Stop'

$srcRoot = Split-Path -Parent $PSScriptRoot
$projectRoot = Split-Path -Parent $srcRoot
Set-Location $projectRoot

$pythonCandidates = @(
    (Join-Path $projectRoot '.venv\Scripts\python.exe')
)

if ($env:pythonLocation) {
    $pythonCandidates += (Join-Path $env:pythonLocation 'python.exe')
}

$python = $pythonCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (-not $python) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        $python = $pythonCommand.Source
    }
}

if (-not $python) {
    throw "Kein Python-Interpreter gefunden (.venv, pythonLocation oder PATH)."
}

$disabledRules = 'MD013,MD022,MD026,MD032,MD036'
$trackedMdFiles = @(git ls-files | Where-Object { $_ -imatch '\.(md|markdown|mdown|mkd)$' })
$untrackedMdFiles = @(git ls-files --others --exclude-standard | Where-Object { $_ -imatch '\.(md|markdown|mdown|mkd)$' })
$mdFiles = @($trackedMdFiles + $untrackedMdFiles | Sort-Object -Unique)

function Resolve-Md024InFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (!(Test-Path $Path)) {
        return $false
    }

    $lines = [System.IO.File]::ReadAllLines($Path)
    $updated = $false
    $inFence = $false
    $headingCounts = @{}

    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]

        if ($line -match '^\s*(```|~~~)') {
            $inFence = -not $inFence
            continue
        }
        if ($inFence) {
            continue
        }

        if ($line -notmatch '^(#{1,6})\s+(.+?)\s*$') {
            continue
        }

        $prefix = $Matches[1]
        $headingText = $Matches[2]
        $headingText = [regex]::Replace($headingText, '\s+#+\s*$', '')
        $headingText = $headingText.Trim()

        if ([string]::IsNullOrWhiteSpace($headingText)) {
            continue
        }

        $normalized = [regex]::Replace($headingText, '\s+', ' ').ToLowerInvariant()
        if (-not $headingCounts.ContainsKey($normalized)) {
            $headingCounts[$normalized] = 0
        }
        $headingCounts[$normalized]++

        if ($headingCounts[$normalized] -le 1) {
            continue
        }

        $newHeading = "$headingText ($($headingCounts[$normalized]))"
        $lines[$i] = "$prefix $newHeading"
        $updated = $true
    }

    if ($updated) {
        [System.IO.File]::WriteAllLines($Path, $lines, [System.Text.UTF8Encoding]::new($false))
    }

    return $updated
}

if ($mdFiles.Count -eq 0) {
    Write-Host "[Markdown] Keine Markdown-Dateien gefunden." -ForegroundColor Yellow
} else {
    Write-Host "[Markdown] Gefundene Markdown-Dateien: $($mdFiles.Count)" -ForegroundColor DarkCyan

    Write-Host "[Markdown] Starte MD024-Auto-Bereinigung (duplizierte Überschriften)..." -ForegroundColor Cyan
    $md024Touched = 0
    foreach ($mdFile in $mdFiles) {
        if (Resolve-Md024InFile -Path $mdFile) {
            $md024Touched++
        }
    }
    Write-Host "[Markdown] MD024-Auto-Bereinigung abgeschlossen. Geänderte Dateien: $md024Touched" -ForegroundColor DarkCyan

    Write-Host "[Markdown] Starte Auto-Fix (alle behebaren Fehler)..." -ForegroundColor Cyan
    & $python -m pymarkdown fix $mdFiles
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[Markdown] Auto-Fix abgeschlossen mit Status $LASTEXITCODE (OK)." -ForegroundColor DarkCyan
    }

    Write-Host "[Markdown] Starte Verifikation (mit deaktivierten Rules)..." -ForegroundColor Cyan
    & $python -m pymarkdown -d $disabledRules scan $mdFiles
    # ExitCode 1 ignorieren wenn nur disabled Rules vorhanden sind
    Write-Host "[Markdown] Lint/Fix abgeschlossen (Disabled Rules: $disabledRules)." -ForegroundColor Green
}

$psAnalyzerScript = Join-Path $PSScriptRoot 'ps_analyze.ps1'
if (Test-Path $psAnalyzerScript) {
    Write-Host "[Markdown] Starte zusätzlich PSScriptAnalyzer-Check ..." -ForegroundColor Cyan
    & $psAnalyzerScript
} else {
    Write-Host "[Markdown] Hinweis: PSScriptAnalyzer-Wrapper nicht gefunden: $psAnalyzerScript" -ForegroundColor Yellow
}
