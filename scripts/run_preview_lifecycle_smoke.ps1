param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [switch]$BuildFirst
)

$ErrorActionPreference = "Stop"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force

if ($BuildFirst) {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "build.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed before preview lifecycle smoke."
    }
    & powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "scripts\sync_docm_code_only.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "DOCM synchronization failed before preview lifecycle smoke."
    }
}

$suiteDocPath = Join-Path $RepoRoot "Practice - Try CleanupSuite Here\CleanupSuite.docm"
$word = $null
$suiteDoc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 1 } catch {}

    $suiteDoc = $word.Documents.Open($suiteDocPath, $false, $true, $false)
    $result = [string]$word.Run("modAlphaSmokeRunner.RunPreviewLifecycleSmokeChecks")
    Write-Host $result
    if ($result -notlike "PASS|*") {
        throw "Preview lifecycle smoke failed: $result"
    }
}
finally {
    if ($suiteDoc -ne $null) {
        try {
            $suiteDoc.ActiveWindow.View.ReadingLayout = $false
            $suiteDoc.ActiveWindow.View.Type = 3
        } catch {}
        $suiteDoc.Close([ref]$false)
    }
    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
}
