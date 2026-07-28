[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$TemplatePath = "",
    [string]$OutputDirectory = "",
    [switch]$PublishRelease
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    throw "Supply the verified v0.9.3-beta-vba-final CleanupSuite.dotm with -TemplatePath."
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = if ($PublishRelease) {
        Join-Path $RepoRoot "release"
    } else {
        Join-Path $RepoRoot "build\v0.9.3-vba-installer"
    }
}

$version = "0.9.3-beta-vba-final"
$tag = "v0.9.3-beta-vba-final"
$expectedTemplateHash =
    "4f815ee0df4bacbda5150a82db2d66739592fa2098184c8d7ec7b535d73ab10c"
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$setupSource = Join-Path $PSScriptRoot "CleanupSuiteForWord.VbaOnly.Setup.cs"

foreach ($required in @($TemplatePath, $compiler, $setupSource)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Required VBA-only installer input was not found: $required"
    }
}

$actualTemplateHash =
    (Get-FileHash -LiteralPath $TemplatePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualTemplateHash -ne $expectedTemplateHash) {
    throw "The supplied template is not the frozen v0.9.3-beta-vba-final release asset."
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$staging = Join-Path $OutputDirectory (".vba-payload-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $staging | Out-Null

try {
    $stagedTemplate = Join-Path $staging "CleanupSuite.dotm"
    $stagedSetup = Join-Path $staging "CleanupSuiteForWord-Setup.exe"
    $stagedManifest = Join-Path $staging "latest-release.ini"
    Copy-Item -LiteralPath $TemplatePath -Destination $stagedTemplate -Force

    & $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Drawing.dll `
        /reference:System.Windows.Forms.dll `
        "/resource:$stagedTemplate,CleanupSuite.dotm" `
        "/out:$stagedSetup" `
        $setupSource
    if ($LASTEXITCODE -ne 0) {
        throw "C# compiler failed with exit code $LASTEXITCODE."
    }

    $setupHash =
        (Get-FileHash -LiteralPath $stagedSetup -Algorithm SHA256).Hash.ToLowerInvariant()
    $manifest = @(
        "# CleanupSuite For Word stable standalone VBA release pointer."
        "version=$version"
        "tag=$tag"
        "template_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/download/$tag/CleanupSuite.dotm"
        "template_sha256=$expectedTemplateHash"
        "setup_url=https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe"
        "setup_sha256=$setupHash"
        "release_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/tag/$tag"
    ) -join "`n"
    [System.IO.File]::WriteAllText(
        $stagedManifest,
        $manifest + "`n",
        (New-Object System.Text.UTF8Encoding($false)))

    $artifacts = @(
        [pscustomobject]@{
            Source = $stagedTemplate
            Destination = (Join-Path $OutputDirectory "CleanupSuite.dotm")
        }
        [pscustomobject]@{
            Source = $stagedSetup
            Destination = (Join-Path $OutputDirectory "CleanupSuiteForWord-Setup.exe")
        }
        [pscustomobject]@{
            Source = $stagedManifest
            Destination = (Join-Path $OutputDirectory "latest-release.ini")
        }
    )

    if ($PublishRelease) {
        $rollback = Join-Path $staging "release-rollback"
        New-Item -ItemType Directory -Force -Path $rollback | Out-Null
        $priorState = @()
        foreach ($artifact in $artifacts) {
            $existed = Test-Path -LiteralPath $artifact.Destination -PathType Leaf
            $backup =
                Join-Path $rollback ([System.IO.Path]::GetFileName($artifact.Destination))
            if ($existed) {
                Copy-Item -LiteralPath $artifact.Destination -Destination $backup -Force
            }
            $priorState += [pscustomobject]@{
                Destination = $artifact.Destination
                Existed = $existed
                Backup = $backup
            }
        }

        try {
            foreach ($artifact in $artifacts) {
                Copy-Item -LiteralPath $artifact.Source `
                    -Destination $artifact.Destination -Force
            }
        }
        catch {
            foreach ($prior in $priorState) {
                if ($prior.Existed) {
                    Copy-Item -LiteralPath $prior.Backup `
                        -Destination $prior.Destination -Force
                } elseif (Test-Path -LiteralPath $prior.Destination -PathType Leaf) {
                    Remove-Item -LiteralPath $prior.Destination -Force
                }
            }
            throw
        }
    } else {
        foreach ($artifact in $artifacts) {
            Copy-Item -LiteralPath $artifact.Source `
                -Destination $artifact.Destination -Force
        }
    }

    Write-Host "PASS|VBA-Only Installer Build|Default Setup embeds v0.9.3-beta-vba-final."
    Write-Host "ARTIFACT|$(Join-Path $OutputDirectory 'CleanupSuiteForWord-Setup.exe')"
    Write-Host "TEMPLATE_SHA256|$actualTemplateHash"
    Write-Host "SETUP_SHA256|$setupHash"
}
finally {
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force
    }
}
