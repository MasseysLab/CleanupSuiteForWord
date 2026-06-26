param(
    [switch]$NoValidate
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $repoRoot
try {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCmd) {
        $pythonCmd = Get-Command py -ErrorAction SilentlyContinue
    }

    if (-not $pythonCmd) {
        throw 'Python was not found on PATH. Install Python 3 and try again.'
    }

    $args = @('assemble.py')
    if ($NoValidate) {
        $args += '--no-validate'
    }

    Write-Host "Building CleanupSuite from $repoRoot"
    & $pythonCmd.Source @args

    if ($LASTEXITCODE -ne 0) {
        throw "Build failed with exit code $LASTEXITCODE"
    }
}
finally {
    Pop-Location
}
