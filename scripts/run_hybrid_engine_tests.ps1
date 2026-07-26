param(
    [string]$Configuration = "Release"
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
$project = Join-Path $repoRoot "engine\CleanupSuite.Engine.ContractTests\CleanupSuite.Engine.ContractTests.csproj"

& $dotnet restore $project --nologo
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine restore failed."
}

& $dotnet build $project -c $Configuration --no-restore --nologo
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine build failed."
}

$validationDirectory = Join-Path $repoRoot "tmp"
New-Item -ItemType Directory -Force -Path $validationDirectory | Out-Null
$resultFixture = Join-Path $validationDirectory "hybrid-engine-result-validation.json"
if (Test-Path -LiteralPath $resultFixture) {
    Remove-Item -LiteralPath $resultFixture -Force
}

& $dotnet run --project $project -c $Configuration --no-build --no-restore -- --result-output $resultFixture
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine contract tests failed."
}

$engineExe = Join-Path $repoRoot "engine\CleanupSuite.Engine\bin\$Configuration\net8.0-windows\CleanupSuite.Engine.exe"
& python (Join-Path $repoRoot "scripts\validate_hybrid_engine_capabilities.py") --engine $engineExe
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine capability validation failed."
}

& python (Join-Path $repoRoot "scripts\validate_hybrid_engine_result.py") --result $resultFixture
if ($LASTEXITCODE -ne 0) {
    throw "Hybrid engine result validation failed."
}

Remove-Item -LiteralPath $resultFixture -Force
