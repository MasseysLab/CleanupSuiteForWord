param(
    [switch]$Quick
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $repoRoot
try {
    Write-Host 'Running CleanupSuite sanity checks...'
    & powershell -ExecutionPolicy Bypass -File .\build.ps1

    if (-not $Quick) {
        Write-Host ''
        Write-Host 'Reminder: test the rebuilt distributable in Word before merging.'
    }
}
finally {
    Pop-Location
}
