[CmdletBinding()]
param(
    [string]$RepoRoot = "",
    [string]$SuiteDotmPath = "",
    [string]$EnginePath = "",
    [string]$WorkRoot = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
if ([string]::IsNullOrWhiteSpace($SuiteDotmPath)) {
    $SuiteDotmPath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP\CleanupSuite.dotm"
}
if ([string]::IsNullOrWhiteSpace($EnginePath)) {
    $EnginePath = Join-Path $RepoRoot "build\hybrid-engine-dev\win-x64\CleanupSuite.Engine.exe"
}
if ([string]::IsNullOrWhiteSpace($WorkRoot)) {
    $WorkRoot = Join-Path $RepoRoot ("build\installed-manifest-smoke-" + [guid]::NewGuid().ToString("N"))
}

$WorkRoot = [System.IO.Path]::GetFullPath($WorkRoot)
$buildRoot = [System.IO.Path]::GetFullPath((Join-Path $RepoRoot "build")).TrimEnd("\") + "\"
if (-not $WorkRoot.StartsWith($buildRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The installed-manifest smoke root must stay under the repository build folder."
}
foreach ($required in @($SuiteDotmPath, $EnginePath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Required installed-manifest smoke input was not found: $required"
    }
}

New-Item -ItemType Directory -Force -Path $WorkRoot | Out-Null
$payloadTemplate = Join-Path $WorkRoot "CleanupSuite.dotm"
$installerOutput = Join-Path $WorkRoot "Setup"
$installationRoot = Join-Path $WorkRoot "Installed"
Copy-Item -LiteralPath $SuiteDotmPath -Destination $payloadTemplate -Force

& powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "scripts\sync_docm_code_only.ps1") `
    -TargetDocPaths $payloadTemplate
if ($LASTEXITCODE -ne 0) {
    throw "The disposable hybrid template could not be synchronized."
}

& powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "installer\build-installer.ps1") `
    -TemplatePath $payloadTemplate `
    -EnginePath $EnginePath `
    -OutputDirectory $installerOutput
if ($LASTEXITCODE -ne 0) {
    throw "The disposable hybrid Setup could not be built."
}

$setupPath = Join-Path $installerOutput "CleanupSuiteForWord-Setup.exe"
$setupArguments = @(
    "/silent",
    "/test-root=`"$installationRoot`"",
    "/install"
)
$setupProcess = Start-Process -FilePath $setupPath -ArgumentList $setupArguments `
    -Wait -PassThru -WindowStyle Hidden
if ($setupProcess.ExitCode -ne 0) {
    throw "The disposable hybrid Setup install failed with exit code $($setupProcess.ExitCode)."
}

$oldAppData = $env:APPDATA
$oldLocalAppData = $env:LOCALAPPDATA
$oldDevEngine = $env:CLEANUPSUITE_HYBRID_DEV_ENGINE
$oldDevHash = $env:CLEANUPSUITE_HYBRID_DEV_SHA256
$oldSuppressAutoExec = $env:CLEANUPSUITE_SUPPRESS_AUTOEXEC
$oldTestJobsRoot = $env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT
$word = $null
$testDocument = $null
try {
    $env:APPDATA = Join-Path $installationRoot "Roaming"
    $env:LOCALAPPDATA = Join-Path $installationRoot "Local"
    Remove-Item Env:CLEANUPSUITE_HYBRID_DEV_ENGINE -ErrorAction SilentlyContinue
    Remove-Item Env:CLEANUPSUITE_HYBRID_DEV_SHA256 -ErrorAction SilentlyContinue
    $env:CLEANUPSUITE_SUPPRESS_AUTOEXEC = "1"
    $env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT = Join-Path $installationRoot "Local\MasseysLab\CleanupSuite\Jobs"
    New-Item -ItemType Directory -Force -Path $env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT | Out-Null

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 1 } catch {}
    $installedTemplate = Join-Path $installationRoot "Roaming\Microsoft\Word\STARTUP\CleanupSuite.dotm"
    $loadedTemplate = $null
    foreach ($addIn in $word.AddIns) {
        if ($addIn.FullName -eq $installedTemplate) {
            $loadedTemplate = $addIn
            break
        }
    }
    if ($null -eq $loadedTemplate) {
        $loadedTemplate = $word.AddIns.Add($installedTemplate, $true)
    } else {
        $loadedTemplate.Installed = $true
    }
    $testDocument = $word.Documents.Add()
    $testDocument.Activate()
    $result = [string]$word.Run(
        "modAlphaSmokeRunner.RunHybridInstalledManifestSmokeCheck")
    Write-Host $result
    if ($result -notlike "PASS|Hybrid Installed Manifest|*") {
        throw "The production installed-manifest bridge smoke failed."
    }
}
finally {
    if ($testDocument -ne $null) {
        try { $testDocument.Close([ref]$false) } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    $env:APPDATA = $oldAppData
    $env:LOCALAPPDATA = $oldLocalAppData
    if ($null -eq $oldDevEngine) {
        Remove-Item Env:CLEANUPSUITE_HYBRID_DEV_ENGINE -ErrorAction SilentlyContinue
    } else {
        $env:CLEANUPSUITE_HYBRID_DEV_ENGINE = $oldDevEngine
    }
    if ($null -eq $oldDevHash) {
        Remove-Item Env:CLEANUPSUITE_HYBRID_DEV_SHA256 -ErrorAction SilentlyContinue
    } else {
        $env:CLEANUPSUITE_HYBRID_DEV_SHA256 = $oldDevHash
    }
    if ($null -eq $oldSuppressAutoExec) {
        Remove-Item Env:CLEANUPSUITE_SUPPRESS_AUTOEXEC -ErrorAction SilentlyContinue
    } else {
        $env:CLEANUPSUITE_SUPPRESS_AUTOEXEC = $oldSuppressAutoExec
    }
    if ($null -eq $oldTestJobsRoot) {
        Remove-Item Env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT -ErrorAction SilentlyContinue
    } else {
        $env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT = $oldTestJobsRoot
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
