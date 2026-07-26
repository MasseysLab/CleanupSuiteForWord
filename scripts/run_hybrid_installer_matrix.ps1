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
    $SetupPath = Join-Path $RepoRoot "build\installer-task8\CleanupSuiteForWord-Setup.exe"
}
if ([string]::IsNullOrWhiteSpace($TestRoot)) {
    $TestRoot = Join-Path $RepoRoot ("build\installer-matrix-" + [guid]::NewGuid().ToString("N"))
}
$SetupPath = (Resolve-Path -LiteralPath $SetupPath).Path
$TestRoot = [System.IO.Path]::GetFullPath($TestRoot)
$buildRoot = [System.IO.Path]::GetFullPath((Join-Path $RepoRoot "build")).TrimEnd("\") + "\"
if (-not $TestRoot.StartsWith($buildRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The installer matrix test root must stay under the repository build folder."
}
New-Item -ItemType Directory -Force -Path $TestRoot | Out-Null
$stateReport = Join-Path $TestRoot "state-report.ini"

function Read-StateReport {
    $result = @{}
    foreach ($line in Get-Content -LiteralPath $stateReport) {
        if ($line -match "^([^=]+)=(.*)$") {
            $result[$matches[1]] = $matches[2]
        }
    }
    return $result
}

function Invoke-SetupAction {
    param(
        [string]$Action,
        [int]$ExpectedExitCode = 0,
        [string[]]$ExtraArguments = @()
    )
    $arguments = @(
        "/silent",
        "/test-root=`"$TestRoot`"",
        "/$Action",
        "/state-report=`"$stateReport`""
    ) + $ExtraArguments
    $process = Start-Process -FilePath $SetupPath -ArgumentList $arguments `
        -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne $ExpectedExitCode) {
        throw "Setup action '$Action' returned $($process.ExitCode), expected $ExpectedExitCode."
    }
    return Read-StateReport
}

function Assert-State {
    param(
        [hashtable]$Report,
        [string]$ExpectedState,
        [string]$ExpectedValid = ""
    )
    if ($Report.state -ne $ExpectedState) {
        throw "Expected installer state '$ExpectedState', received '$($Report.state)': $($Report.detail)"
    }
    if ($ExpectedValid -and $Report.package_valid -ne $ExpectedValid) {
        throw "Expected package_valid=$ExpectedValid, received $($Report.package_valid)."
    }
}

$report = Invoke-SetupAction "inspect"
Assert-State $report "Absent" "true"

$productFolder = Join-Path $TestRoot "Local\MasseysLab\CleanupSuiteForWord"
$startupFolder = Join-Path $TestRoot "Roaming\Microsoft\Word\STARTUP"
$templatePath = Join-Path $startupFolder "CleanupSuite.dotm"
$enginePath = Join-Path $productFolder "Engine\CleanupSuite.Engine.exe"
$protocolPath = Join-Path $productFolder "Contracts\Hybrid\v1\protocol.json"
$operationsPath = Join-Path $productFolder "Contracts\Hybrid\v1\operation-vocabulary.json"
$manifestPath = Join-Path $productFolder "installation-manifest.json"
$backupFolder = Join-Path $productFolder "Backups"
$settingsSentinel = Join-Path $productFolder "UserSettings.keep"
New-Item -ItemType Directory -Force -Path $productFolder | Out-Null
Set-Content -LiteralPath $settingsSentinel -Value "preserve-user-settings"

$report = Invoke-SetupAction "install"
Assert-State $report "CurrentHealthy" "true"
foreach ($required in @($templatePath, $enginePath, $protocolPath, $operationsPath, $manifestPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Install did not create matched component: $required"
    }
}

[System.IO.File]::AppendAllText($enginePath, "intentional matrix damage")
$report = Invoke-SetupAction "inspect"
Assert-State $report "CurrentDamaged" "false"
if ($report.detail -notmatch "Repair") {
    throw "A damaged matched installation did not recommend Repair."
}

$damagedEngineHash = (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash
$report = Invoke-SetupAction "repair" 1 @("/test-fail-after=engine")
Assert-State $report "CurrentDamaged" "false"
if ((Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash -ne $damagedEngineHash) {
    throw "A failed Repair did not roll the damaged pre-repair engine back exactly."
}
if (@(Get-ChildItem -LiteralPath $productFolder -Directory -Filter ".staging-*").Count -ne 0) {
    throw "A failed Repair left temporary payload files behind."
}

$report = Invoke-SetupAction "repair"
Assert-State $report "CurrentHealthy" "true"
if (-not (Test-Path -LiteralPath $settingsSentinel)) {
    throw "Repair removed the user-settings sentinel."
}

1..4 | ForEach-Object {
    $repairReport = Invoke-SetupAction "repair"
    Assert-State $repairReport "CurrentHealthy" "true"
}
$backupCount = @(Get-ChildItem -LiteralPath $backupFolder -Filter "CleanupSuite-*.dotm").Count
if ($backupCount -ne 3) {
    throw "Expected exactly 3 retained backups after repeated repair, found $backupCount."
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$manifest.suiteVersion = "0.9.4-beta"
[System.IO.File]::WriteAllText(
    $manifestPath,
    (($manifest | ConvertTo-Json -Depth 6) + "`n"),
    (New-Object System.Text.UTF8Encoding($false)))
$report = Invoke-SetupAction "inspect"
Assert-State $report "Older" "false"
$report = Invoke-SetupAction "update"
Assert-State $report "CurrentHealthy" "true"

$report = Invoke-SetupAction "uninstall"
Assert-State $report "Absent" "true"
foreach ($removed in @($templatePath, $enginePath, $protocolPath, $operationsPath, $manifestPath)) {
    if (Test-Path -LiteralPath $removed) {
        throw "Uninstall left a product component behind: $removed"
    }
}
if (-not (Test-Path -LiteralPath $settingsSentinel)) {
    throw "Uninstall removed the user-settings sentinel."
}
if (@(Get-ChildItem -LiteralPath $backupFolder -Filter "CleanupSuite-*.dotm").Count -ne 3) {
    throw "Uninstall did not preserve the newest three backups."
}

New-Item -ItemType Directory -Force -Path $startupFolder | Out-Null
Copy-Item -LiteralPath (Join-Path $RepoRoot "release\CleanupSuite.dotm") `
    -Destination $templatePath -Force
$report = Invoke-SetupAction "inspect"
Assert-State $report "Older" "false"
$report = Invoke-SetupAction "update"
Assert-State $report "CurrentHealthy" "true"

Write-Host "PASS|Hybrid Installer Matrix|Absent, legacy, older, current, damaged, repair, update, and uninstall states passed in isolation."
Write-Host "PASS|Hybrid Installer Safety|Matched hashes, rollback-ready writes, three backups, settings preservation, and payload cleanup passed."
Write-Host "TEST_ROOT|$TestRoot"
