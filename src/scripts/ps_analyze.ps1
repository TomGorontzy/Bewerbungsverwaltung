param(
    [string]$Path,
    [string]$SettingsPath
)

$ErrorActionPreference = 'Stop'

$srcRoot = Split-Path -Parent $PSScriptRoot
$projectRoot = Split-Path -Parent $srcRoot
if ([string]::IsNullOrWhiteSpace($Path)) {
    $Path = $projectRoot
}
if ([string]::IsNullOrWhiteSpace($SettingsPath)) {
    $SettingsPath = Join-Path $srcRoot 'PSScriptAnalyzerSettings.psd1'
}

if (!(Test-Path $SettingsPath)) {
    throw "PSScriptAnalyzer-Settings nicht gefunden: $SettingsPath"
}

Import-Module PSScriptAnalyzer -ErrorAction Stop

Write-Host "[PSSA] Analyse gestartet: $Path" -ForegroundColor Cyan
$results = Invoke-ScriptAnalyzer -Path $Path -Recurse -Settings $SettingsPath

if (-not $results -or $results.Count -eq 0) {
    Write-Host "[PSSA] Keine Findings." -ForegroundColor Green
    return
}

$errors = @($results | Where-Object { $_.Severity -eq 'Error' })
$warnings = @($results | Where-Object { $_.Severity -eq 'Warning' })
$info = @($results | Where-Object { $_.Severity -eq 'Information' })

Write-Host "[PSSA] Findings: Errors=$($errors.Count), Warnings=$($warnings.Count), Info=$($info.Count)" -ForegroundColor Yellow

$results |
    Select-Object Severity, RuleName, ScriptName, Line, Message |
    Sort-Object Severity, ScriptName, Line |
    Format-Table -AutoSize |
    Out-String -Width 260 |
    Write-Host

if ($errors.Count -gt 0 -or $warnings.Count -gt 0) {
    throw "PSScriptAnalyzer meldet Error/Warning. Bitte vor dem Fortfahren beheben oder Rules gezielt konfigurieren."
}

Write-Host "[PSSA] Nur Information-Findings vorhanden." -ForegroundColor DarkCyan
