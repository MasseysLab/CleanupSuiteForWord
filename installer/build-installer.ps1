[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$TemplatePath = "$env:APPDATA\Microsoft\Word\STARTUP\CleanupSuite.dotm"
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
$version = "0.9.2-alpha"
$tag = "v0.9.2-alpha"
$releaseDirectory = Join-Path $RepoRoot "release"
$releaseTemplate = Join-Path $releaseDirectory "CleanupSuite.dotm"
$setupSource = Join-Path $PSScriptRoot "CleanupSuiteForWord.Setup.cs"
$setupOutput = Join-Path $releaseDirectory "CleanupSuiteForWord-Setup.exe"
$manifestPath = Join-Path $releaseDirectory "latest-release.ini"
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "The release template was not found: $TemplatePath"
}
if (-not (Test-Path -LiteralPath $compiler)) {
    throw "The Windows C# compiler was not found: $compiler"
}

New-Item -ItemType Directory -Force -Path $releaseDirectory | Out-Null
Copy-Item -LiteralPath $TemplatePath -Destination $releaseTemplate -Force

& $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
    /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll `
    "/resource:$releaseTemplate,CleanupSuite.dotm" "/out:$setupOutput" $setupSource
if ($LASTEXITCODE -ne 0) { throw "C# compiler failed with exit code $LASTEXITCODE." }

$templateHash = (Get-FileHash -LiteralPath $releaseTemplate -Algorithm SHA256).Hash.ToLowerInvariant()
$manifest = @(
    "# CleanupSuite For Word release pointer. This file is intentionally stable on main."
    "version=$version"
    "tag=$tag"
    "template_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/download/$tag/CleanupSuite.dotm"
    "template_sha256=$templateHash"
    "setup_url=https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe"
    "release_url=https://github.com/MasseysLab/CleanupSuiteForWord/releases/tag/$tag"
) -join "`n"
[System.IO.File]::WriteAllText($manifestPath, $manifest + "`n", (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Built $setupOutput"
Write-Host "Embedded template SHA-256: $templateHash"
