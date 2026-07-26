[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$CandidateDirectory = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
if ([string]::IsNullOrWhiteSpace($CandidateDirectory)) {
    $CandidateDirectory = Join-Path $RepoRoot "build\installer-signed-candidate"
}

$candidateFullPath = [System.IO.Path]::GetFullPath($CandidateDirectory)
$buildRoot = [System.IO.Path]::GetFullPath((Join-Path $RepoRoot "build")).TrimEnd("\") + "\"
if (-not $candidateFullPath.StartsWith($buildRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The signed candidate must stay under the repository build folder."
}

$setupPath = Join-Path $candidateFullPath "CleanupSuiteForWord-Setup.exe"
$payloadRoot = Join-Path $candidateFullPath "candidate-payload"
$manifestPath = Join-Path $payloadRoot "installation-manifest.json"
foreach ($required in @($setupPath, $payloadRoot, $manifestPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Required signed-candidate artifact was not found: $required"
    }
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ($manifest.manifestVersion -ne "1.0" -or
    $manifest.suiteVersion -ne "0.9.5-beta" -or
    $manifest.protocolVersion -ne "1.0" -or
    $manifest.publisher -ne "MasseysLab" -or
    $manifest.authenticodeRequired -ne $true) {
    throw "The signed candidate manifest does not match the approved 0.9.5-beta contract."
}

$approved = @{
    "word-template" = "CleanupSuite.dotm"
    "analysis-engine" = "Engine\CleanupSuite.Engine.exe"
    "tool-definitions" = "Contracts\Hybrid\v1\protocol.json"
    "rules" = "Contracts\Hybrid\v1\operation-vocabulary.json"
}
$approvedVersions = @{
    "word-template" = "0.9.5-beta"
    "analysis-engine" = "0.2.0"
    "tool-definitions" = "1.0.0"
    "rules" = "1.0.0"
}
if (@($manifest.components).Count -ne $approved.Count) {
    throw "The signed candidate manifest must contain exactly four approved components."
}

$seen = @{}
foreach ($component in @($manifest.components)) {
    $id = [string]$component.id
    if (-not $approved.ContainsKey($id) -or $seen.ContainsKey($id)) {
        throw "The signed candidate contains an unknown or duplicate component: $id"
    }
    $seen[$id] = $true
    $expectedRelativePath = $approved[$id]
    if ([string]$component.relativePath -ne $expectedRelativePath) {
        throw "The signed candidate component path is not approved: $id"
    }
    if ([string]$component.version -ne $approvedVersions[$id]) {
        throw "The signed candidate component version is not approved: $id"
    }

    $componentPath = Join-Path $payloadRoot $expectedRelativePath
    if (-not (Test-Path -LiteralPath $componentPath -PathType Leaf)) {
        throw "The signed candidate component is missing: $id"
    }
    $item = Get-Item -LiteralPath $componentPath
    $actualHash = (Get-FileHash -LiteralPath $componentPath -Algorithm SHA256).Hash
    if ($item.Length -ne [long]$component.byteLength -or
        $actualHash -ne ([string]$component.sha256).ToUpperInvariant()) {
        throw "The signed candidate component does not match its manifest: $id"
    }
}
if ($seen.Count -ne $approved.Count) {
    throw "The signed candidate is missing an approved component."
}

$enginePath = Join-Path $payloadRoot $approved["analysis-engine"]
foreach ($signedBinary in @($enginePath, $setupPath)) {
    $signature = Get-AuthenticodeSignature -LiteralPath $signedBinary
    if ($signature.Status -ne "Valid") {
        throw "Authenticode validation failed for $signedBinary`: $($signature.Status)"
    }
}

$defenderStatus = Get-MpComputerStatus
if (-not $defenderStatus.AntivirusEnabled -or
    -not $defenderStatus.RealTimeProtectionEnabled) {
    throw "Microsoft Defender antivirus and real-time protection must be enabled for this gate."
}

$scanner = "$env:ProgramFiles\Windows Defender\MpCmdRun.exe"
if (-not (Test-Path -LiteralPath $scanner)) {
    $platformFolder = Get-ChildItem `
        "$env:ProgramData\Microsoft\Windows Defender\Platform" -Directory |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if ($null -ne $platformFolder) {
        $scanner = Join-Path $platformFolder.FullName "MpCmdRun.exe"
    }
}
if (-not (Test-Path -LiteralPath $scanner)) {
    throw "Microsoft Defender command-line scanner was not found."
}

foreach ($scanTarget in @($enginePath, $setupPath)) {
    & $scanner -Scan -ScanType 3 -File $scanTarget -DisableRemediation
    if ($LASTEXITCODE -ne 0) {
        throw "Microsoft Defender did not pass $scanTarget (exit code $LASTEXITCODE)."
    }
}

Write-Host "PASS|Beta Release Candidate Manifest|All four matched components passed path, length, and SHA-256 validation."
Write-Host "PASS|Beta Release Candidate Signatures|The engine and Setup have valid Authenticode signatures."
Write-Host "PASS|Beta Release Candidate Defender|The exact signed engine and Setup reported no threats."
