# Start-AI-Session.ps1
# Run this at the start of a local Codex/ChatGPT editing session.
# It creates a timestamped project backup, checks that the backup does not include backups/, and prints Git state.

$Project = "C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word"
$BackupDir = Join-Path $Project "backups"

if (!(Test-Path $Project)) {
    throw "Project folder not found: $Project"
}

New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

$Stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ZipPath = Join-Path $BackupDir "CleanupSuite_backup_$Stamp.zip"
$Temp = Join-Path $env:TEMP "CleanupSuite_backup_staging_$Stamp"

if (Test-Path $Temp) {
    Remove-Item -Path $Temp -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $Temp | Out-Null

Get-ChildItem -Path $Project -Force |
    Where-Object { $_.Name -ne "backups" } |
    Copy-Item -Destination $Temp -Recurse -Force

Compress-Archive -Path "$Temp\*" -DestinationPath $ZipPath -Force
Remove-Item -Path $Temp -Recurse -Force

Write-Host ""
Write-Host "Backup created:"
Write-Host $ZipPath

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $Zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    $HasBackupsFolder = $false
    foreach ($Entry in $Zip.Entries) {
        if ($Entry.FullName -like "backups/*" -or $Entry.FullName -like "backups\*") {
            $HasBackupsFolder = $true
            break
        }
    }
    $Zip.Dispose()
    if ($HasBackupsFolder) {
        Write-Warning "Backup verification failed: zip appears to contain backups/."
    } else {
        Write-Host "Backup verification passed: backups/ was not included."
    }
} catch {
    Write-Warning "Could not inspect zip contents for backups/ exclusion: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Before editing, read:"
Write-Host (Join-Path $Project "AI_HANDOFF.md")
Write-Host (Join-Path $Project "AI_RULES.md")
Write-Host (Join-Path $Project "CODEX_SETUP_TASK.md")
Write-Host (Join-Path $Project "REPO_FINDINGS_AND_SYNC_CHECK.md")

Write-Host ""
Write-Host "Next known technical task, if Chris wants it handled:"
Write-Host (Join-Path $Project "PREVIEW_ACTIONS_NEXT_TASK.md")

Write-Host ""
Write-Host "Git status snapshot:"
Push-Location $Project
try {
    git remote -v
    git branch --show-current
    git status --short
} catch {
    Write-Warning "Could not run Git commands: $($_.Exception.Message)"
}
Pop-Location
Write-Host ""
