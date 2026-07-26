param(
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$candidateDotnet = @(
    $env:CLEANUPSUITE_DOTNET,
    (Join-Path $env:USERPROFILE ".dotnet-cleanupsuite-dev\dotnet.exe"),
    (Get-Command dotnet -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue)
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $candidateDotnet) {
    throw "A .NET 8 SDK is required. Set CLEANUPSUITE_DOTNET or install the SDK."
}

$dotnet = $candidateDotnet[0]
$sdkList = & $dotnet --list-sdks
if ($LASTEXITCODE -ne 0 -or -not ($sdkList | Select-String -Pattern '^8\.')) {
    throw "The selected dotnet host does not provide a .NET 8 SDK: $dotnet"
}

$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:DOTNET_NOLOGO = "1"
$project = Join-Path $repoRoot "engine\CleanupSuite.Engine\CleanupSuite.Engine.csproj"
$publishDirectory = Join-Path $repoRoot "build\hybrid-engine-dev\$RuntimeIdentifier"

& $dotnet publish $project `
    -c $Configuration `
    -r $RuntimeIdentifier `
    --self-contained true `
    --nologo `
    -o $publishDirectory `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:PublishTrimmed=false `
    -p:DebugType=None `
    -p:DebugSymbols=false
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine self-contained publication failed."
}

$engine = Join-Path $publishDirectory "CleanupSuite.Engine.exe"
& python (Join-Path $repoRoot "scripts\validate_hybrid_engine_capabilities.py") --engine $engine
if ($LASTEXITCODE -ne 0) {
    throw "Published hybrid engine capability validation failed."
}

$hash = Get-FileHash -Algorithm SHA256 -LiteralPath $engine
Write-Output "PASS|Hybrid Engine Publish|Self-contained $RuntimeIdentifier engine created."
Write-Output "ARTIFACT|$engine"
Write-Output "SHA256|$($hash.Hash)"
