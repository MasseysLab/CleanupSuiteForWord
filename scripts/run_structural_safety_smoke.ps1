param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [switch]$SkipSync,
    [switch]$SkipFixtureExport
)

$ErrorActionPreference = "Stop"

if (-not $SkipSync) {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "scripts\sync_docm_code_only.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "CleanupSuite source synchronization failed."
    }
}

$suitePath = Join-Path $RepoRoot "Practice - Try CleanupSuite Here\CleanupSuite.docm"
$fixturePath = Join-Path $RepoRoot "Testing Files\Regression - Structural Safety.docx"
$reportDir = Join-Path $RepoRoot "docs\test-results"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportPath = Join-Path $reportDir "structural_safety_smoke_$timestamp.txt"
$pdfPath = Join-Path $reportDir "structural_safety_fixture_$timestamp.pdf"

New-Item -ItemType Directory -Force -Path $reportDir | Out-Null

$word = $null
$suite = $null
$fixture = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 1 } catch {}

    $suite = $word.Documents.Open($suitePath, $false, $true, $false)
    $macroName = "modAlphaSmokeRunner.RunStructuralSafetySmokeCheck"
    $result = [string]$word.Run($macroName)
    if (-not $result.StartsWith("PASS|Structural Safety|")) {
        throw "Word structural smoke failed: $result"
    }

    $reportLines = @($result)
    $fixture = $word.Documents.Open($fixturePath, $false, $true, $false)
    $fixtureResult = [string]$word.Run("modAlphaSmokeRunner.RunStructuralFixtureSmokeCheck")
    if (-not $fixtureResult.StartsWith("PASS|Structural Fixture|")) {
        throw "Word structural fixture smoke failed: $fixtureResult"
    }
    $reportLines += $fixtureResult
    if (-not $SkipFixtureExport) {
        $fixture.ExportAsFixedFormat($pdfPath, 17)
        $reportLines += "PASS|Fixture PDF|Word rendered the structural-safety fixture to $pdfPath"
    }
    $fixture.Close([ref]$false)
    $fixture = $null
    $reportLines | Set-Content -LiteralPath $reportPath -Encoding utf8

    Write-Output $result
    Write-Output $fixtureResult
    if (-not $SkipFixtureExport) {
        Write-Output "PASS|Fixture PDF|$pdfPath"
    }
    Write-Output "REPORT|$reportPath"
}
finally {
    if ($fixture -ne $null) {
        $fixture.Close([ref]$false)
    }
    if ($suite -ne $null) {
        $suite.Close([ref]$false)
    }
    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
