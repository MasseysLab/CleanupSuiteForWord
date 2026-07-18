param(
    [switch]$Quick
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $repoRoot
try {
    Write-Host 'Running CleanupSuite sanity checks...'
    & powershell -ExecutionPolicy Bypass -File .\build.ps1

    Write-Host ''
    Write-Host 'Running regression tests...'
    & python -m unittest discover -s tests
    if ($LASTEXITCODE -ne 0) {
        throw 'Regression tests failed.'
    }

    if (-not $Quick) {
        Write-Host ''
        Write-Host 'Reminder: test the rebuilt distributable in Word before merging.'
Write-Host 'Run the Word alpha smoke runner with powershell -ExecutionPolicy Bypass -File .\scripts\run_word_alpha_smoke.ps1 for the 0.9.2-alpha release gate.'
    }
}
finally {
    Pop-Location
}
