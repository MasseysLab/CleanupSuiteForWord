[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$SetupPath = "",
    [string]$TestRoot = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
if ([string]::IsNullOrWhiteSpace($SetupPath)) {
    $SetupPath =
        Join-Path $RepoRoot "build\v0.9.3-vba-installer\CleanupSuiteForWord-Setup.exe"
}
if ([string]::IsNullOrWhiteSpace($TestRoot)) {
    $TestRoot =
        Join-Path $RepoRoot ("build\vba-installer-matrix-" + [guid]::NewGuid().ToString("N"))
}

$TestRoot = [System.IO.Path]::GetFullPath($TestRoot)
$buildRoot =
    [System.IO.Path]::GetFullPath((Join-Path $RepoRoot "build")).TrimEnd("\") + "\"
if (-not $TestRoot.StartsWith(
        $buildRoot,
        [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The VBA-only installer matrix root must stay under the repository build folder."
}
if (-not (Test-Path -LiteralPath $SetupPath -PathType Leaf)) {
    throw "The VBA-only Setup was not found: $SetupPath"
}

$expectedHash =
    "4F815EE0DF4BACBDA5150A82DB2D66739592FA2098184C8D7EC7B535D73AB10C"
$startupFolder = Join-Path $TestRoot "Roaming\Microsoft\Word\STARTUP"
$productFolder = Join-Path $TestRoot "Local\MasseysLab\CleanupSuiteForWord"
$backupFolder = Join-Path $productFolder "Backups"
$templatePath = Join-Path $startupFolder "CleanupSuite.dotm"
$cachedSetup = Join-Path $productFolder "CleanupSuiteForWord-Setup.exe"

function Invoke-Setup {
    param([string[]]$Arguments)

    $allArguments = @(
        "/silent"
        "/test-root=`"$TestRoot`""
    ) + $Arguments
    $process = Start-Process -FilePath $SetupPath -ArgumentList $allArguments `
        -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "VBA-only Setup failed with exit code $($process.ExitCode): $($Arguments -join ' ')"
    }
}

function Assert-InstalledTemplate {
    if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) {
        throw "VBA-only Setup did not install CleanupSuite.dotm."
    }
    $actual = (Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash
    if ($actual -ne $expectedHash) {
        throw "Installed template hash did not match v0.9.3-beta-vba-final."
    }
}

New-Item -ItemType Directory -Force -Path $startupFolder | Out-Null
[System.IO.File]::WriteAllBytes(
    $templatePath,
    [System.Text.Encoding]::UTF8.GetBytes("legacy-template-sentinel"))

Invoke-Setup @()
Assert-InstalledTemplate
if (-not (Test-Path -LiteralPath $cachedSetup -PathType Leaf)) {
    throw "VBA-only Setup did not cache its uninstall executable."
}
if ((Get-FileHash -LiteralPath $cachedSetup -Algorithm SHA256).Hash -ne
    (Get-FileHash -LiteralPath $SetupPath -Algorithm SHA256).Hash) {
    throw "Cached VBA-only Setup did not match the executable that installed it."
}
$backups = @(Get-ChildItem -LiteralPath $backupFolder -Filter "CleanupSuite-*.dotm")
if ($backups.Count -ne 1) {
    throw "Initial install did not create exactly one legacy-template backup."
}
$legacyBytes = [System.IO.File]::ReadAllBytes($backups[0].FullName)
if ([System.Text.Encoding]::UTF8.GetString($legacyBytes) -ne
    "legacy-template-sentinel") {
    throw "Initial backup did not preserve the prior template exactly."
}

[System.IO.File]::WriteAllBytes(
    $templatePath,
    [System.Text.Encoding]::UTF8.GetBytes("damaged-template-sentinel"))
Invoke-Setup @()
Assert-InstalledTemplate
$backups = @(Get-ChildItem -LiteralPath $backupFolder -Filter "CleanupSuite-*.dotm")
if ($backups.Count -ne 2) {
    throw "Reinstall did not retain the expected second backup."
}

Invoke-Setup @("/uninstall")
if (Test-Path -LiteralPath $templatePath) {
    throw "VBA-only uninstall left the Startup template installed."
}
$backups = @(Get-ChildItem -LiteralPath $backupFolder -Filter "CleanupSuite-*.dotm")
if ($backups.Count -ne 3) {
    throw "VBA-only uninstall did not preserve exactly three newest backups."
}

Write-Host "PASS|VBA-Only Installer Matrix|Install, exact backup, reinstall, hash verification, and uninstall passed."
Write-Host "PASS|VBA-Only Stable Version|The installed template matched the 0.9.3 stable release."
Write-Host "TEST_ROOT|$TestRoot"
