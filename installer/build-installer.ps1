[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$TemplatePath = "",
    [string]$EnginePath = "",
    [string]$OutputDirectory = "",
    [string]$Version = "0.9.5-beta",
    [string]$Tag = "v0.9.5-beta",
    [switch]$SignBuild,
    [switch]$PublishRelease,
    [string]$SigningCertificateThumbprint = "",
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
if ($Version -ne "0.9.5-beta") {
    throw "The current Setup source is bound to 0.9.5-beta. Update the source and tests before changing -Version."
}
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Join-Path $RepoRoot "release\CleanupSuite.dotm"
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = if ($PublishRelease) {
        Join-Path $RepoRoot "release"
    } elseif ($SignBuild) {
        Join-Path $RepoRoot "build\installer-signed-candidate"
    } else {
        Join-Path $RepoRoot "build\installer-dev"
    }
}
$shouldSign = $SignBuild -or $PublishRelease

$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$webExtensions = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\System.Web.Extensions.dll"
$setupSource = Join-Path $PSScriptRoot "CleanupSuiteForWord.Setup.cs"
$protocolSource = Join-Path $RepoRoot "contracts\hybrid\v1\protocol.json"
$operationsSource = Join-Path $RepoRoot "contracts\hybrid\v1\operation-vocabulary.json"

foreach ($required in @($compiler, $webExtensions, $setupSource, $TemplatePath, $protocolSource, $operationsSource)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Required installer input was not found: $required"
    }
}

if ($shouldSign -and [string]::IsNullOrWhiteSpace($SigningCertificateThumbprint)) {
    throw "Signed hybrid builds require -SigningCertificateThumbprint."
}

if ([string]::IsNullOrWhiteSpace($EnginePath)) {
    & powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "scripts\publish_hybrid_engine_dev.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "The self-contained hybrid engine could not be published."
    }
    $EnginePath = Join-Path $RepoRoot "build\hybrid-engine-dev\win-x64\CleanupSuite.Engine.exe"
}
if (-not (Test-Path -LiteralPath $EnginePath)) {
    throw "The self-contained hybrid engine was not found: $EnginePath"
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$staging = Join-Path $OutputDirectory (".payload-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $staging | Out-Null

try {
    $templatePayload = Join-Path $staging "CleanupSuite.dotm"
    $enginePayload = Join-Path $staging "CleanupSuite.Engine.exe"
    $protocolPayload = Join-Path $staging "protocol.json"
    $operationsPayload = Join-Path $staging "operation-vocabulary.json"
    $packageManifest = Join-Path $staging "installation-manifest.json"
    Copy-Item -LiteralPath $TemplatePath -Destination $templatePayload -Force
    Copy-Item -LiteralPath $EnginePath -Destination $enginePayload -Force
    Copy-Item -LiteralPath $protocolSource -Destination $protocolPayload -Force
    Copy-Item -LiteralPath $operationsSource -Destination $operationsPayload -Force

    if ($shouldSign) {
        $certificate = Get-ChildItem "Cert:\CurrentUser\My\$SigningCertificateThumbprint" -ErrorAction Stop
        $engineSignature = Set-AuthenticodeSignature -LiteralPath $enginePayload `
            -Certificate $certificate -TimestampServer $TimestampServer -HashAlgorithm SHA256
        if ($engineSignature.Status -ne "Valid") {
            throw "The staged hybrid engine did not receive a valid Authenticode signature: $($engineSignature.Status)"
        }
    }

    function New-Component {
        param(
            [string]$Id,
            [string]$ComponentVersion,
            [string]$RelativePath,
            [string]$PayloadPath
        )
        $item = Get-Item -LiteralPath $PayloadPath
        [ordered]@{
            id = $Id
            version = $ComponentVersion
            relativePath = $RelativePath
            sha256 = (Get-FileHash -LiteralPath $PayloadPath -Algorithm SHA256).Hash.ToLowerInvariant()
            byteLength = $item.Length
        }
    }

    $manifestObject = [ordered]@{
        manifestVersion = "1.0"
        suiteVersion = $Version
        protocolVersion = "1.0"
        publisher = "MasseysLab"
        authenticodeRequired = $true
        components = @(
            (New-Component "word-template" $Version "CleanupSuite.dotm" $templatePayload)
            (New-Component "analysis-engine" "0.2.0" "Engine\CleanupSuite.Engine.exe" $enginePayload)
            (New-Component "tool-definitions" "1.0.0" "Contracts\Hybrid\v1\protocol.json" $protocolPayload)
            (New-Component "rules" "1.0.0" "Contracts\Hybrid\v1\operation-vocabulary.json" $operationsPayload)
        )
    }
    $manifestJson = $manifestObject | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText(
        $packageManifest,
        $manifestJson + "`n",
        (New-Object System.Text.UTF8Encoding($false)))

    $setupOutput = Join-Path $OutputDirectory "CleanupSuiteForWord-Setup.exe"
    $setupBuildOutput = if ($PublishRelease) {
        Join-Path $staging "CleanupSuiteForWord-Setup.exe"
    } else {
        $setupOutput
    }
    & $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Drawing.dll `
        /reference:System.Windows.Forms.dll `
        "/reference:$webExtensions" `
        "/resource:$packageManifest,CleanupSuite.Payload.Manifest.json" `
        "/resource:$templatePayload,CleanupSuite.Payload.Template.dotm" `
        "/resource:$enginePayload,CleanupSuite.Payload.Engine.exe" `
        "/resource:$protocolPayload,CleanupSuite.Payload.Protocol.json" `
        "/resource:$operationsPayload,CleanupSuite.Payload.Operations.json" `
        "/out:$setupBuildOutput" `
        $setupSource
    if ($LASTEXITCODE -ne 0) {
        throw "C# compiler failed with exit code $LASTEXITCODE."
    }

    if ($shouldSign) {
        $setupSignature = Set-AuthenticodeSignature -LiteralPath $setupBuildOutput `
            -Certificate $certificate -TimestampServer $TimestampServer -HashAlgorithm SHA256
        if ($setupSignature.Status -ne "Valid") {
            throw "Setup did not receive a valid Authenticode signature: $($setupSignature.Status)"
        }
    }

    if ($PublishRelease) {
        $setupHash = (Get-FileHash -LiteralPath $setupBuildOutput -Algorithm SHA256).Hash.ToLowerInvariant()
        $templateHash = (Get-FileHash -LiteralPath $templatePayload -Algorithm SHA256).Hash.ToLowerInvariant()
        $releaseManifest = @(
            "# CleanupSuite For Word release pointer."
            "version=$Version"
            "tag=$Tag"
            "template_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/download/$Tag/CleanupSuite.dotm"
            "template_sha256=$templateHash"
            "setup_url=https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe"
            "setup_sha256=$setupHash"
            "release_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/tag/$Tag"
        ) -join "`n"
        [System.IO.File]::WriteAllText(
            (Join-Path $staging "latest-release.ini"),
            $releaseManifest + "`n",
            (New-Object System.Text.UTF8Encoding($false)))

        $releaseArtifacts = @(
            [pscustomobject]@{
                Source = $templatePayload
                Destination = (Join-Path $OutputDirectory "CleanupSuite.dotm")
            }
            [pscustomobject]@{
                Source = $setupBuildOutput
                Destination = $setupOutput
            }
            [pscustomobject]@{
                Source = (Join-Path $staging "latest-release.ini")
                Destination = (Join-Path $OutputDirectory "latest-release.ini")
            }
        )
        $releaseRollback = Join-Path $staging "release-rollback"
        New-Item -ItemType Directory -Force -Path $releaseRollback | Out-Null
        $releaseState = @()
        foreach ($artifact in $releaseArtifacts) {
            $existed = Test-Path -LiteralPath $artifact.Destination -PathType Leaf
            $backupPath = Join-Path $releaseRollback ([System.IO.Path]::GetFileName($artifact.Destination))
            if ($existed) {
                Copy-Item -LiteralPath $artifact.Destination -Destination $backupPath -Force
            }
            $releaseState += [pscustomobject]@{
                Destination = $artifact.Destination
                Existed = $existed
                BackupPath = $backupPath
            }
        }
        try {
            foreach ($artifact in $releaseArtifacts) {
                Copy-Item -LiteralPath $artifact.Source -Destination $artifact.Destination -Force
            }
        }
        catch {
            foreach ($state in $releaseState) {
                if ($state.Existed) {
                    Copy-Item -LiteralPath $state.BackupPath -Destination $state.Destination -Force
                } elseif (Test-Path -LiteralPath $state.Destination -PathType Leaf) {
                    Remove-Item -LiteralPath $state.Destination -Force
                }
            }
            throw
        }
    } else {
        Copy-Item -LiteralPath $packageManifest `
            -Destination (Join-Path $OutputDirectory "installation-manifest.json") -Force
        if ($SignBuild) {
            $candidateRoot = Join-Path $OutputDirectory "candidate-payload"
            $candidateEngineFolder = Join-Path $candidateRoot "Engine"
            $candidateContractFolder = Join-Path $candidateRoot "Contracts\Hybrid\v1"
            New-Item -ItemType Directory -Force -Path $candidateEngineFolder | Out-Null
            New-Item -ItemType Directory -Force -Path $candidateContractFolder | Out-Null
            Copy-Item -LiteralPath $templatePayload `
                -Destination (Join-Path $candidateRoot "CleanupSuite.dotm") -Force
            Copy-Item -LiteralPath $enginePayload `
                -Destination (Join-Path $candidateEngineFolder "CleanupSuite.Engine.exe") -Force
            Copy-Item -LiteralPath $protocolPayload `
                -Destination (Join-Path $candidateContractFolder "protocol.json") -Force
            Copy-Item -LiteralPath $operationsPayload `
                -Destination (Join-Path $candidateContractFolder "operation-vocabulary.json") -Force
            Copy-Item -LiteralPath $packageManifest `
                -Destination (Join-Path $candidateRoot "installation-manifest.json") -Force
        }
    }

    Write-Host "PASS|Hybrid Installer Build|State-aware Setup built with matched embedded payload."
    Write-Host "ARTIFACT|$setupOutput"
    Write-Host "SHA256|$((Get-FileHash -LiteralPath $setupOutput -Algorithm SHA256).Hash)"
}
finally {
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force
    }
}
