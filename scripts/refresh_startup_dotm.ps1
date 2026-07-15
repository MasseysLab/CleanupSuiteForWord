param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$sourceDocmPath = Join-Path $RepoRoot "CleanupSuite.docm"
if (-not (Test-Path -LiteralPath $sourceDocmPath)) {
    throw "Source docm not found: $sourceDocmPath"
}

$StartupDir = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
$startupDotmPath = Join-Path $StartupDir "CleanupSuite.dotm"
$BackupDir = Join-Path (Split-Path $RepoRoot -Parent) "CleanupSuite external backups\startup-dotm"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$documentContentType = "application/vnd.ms-word.document.macroEnabled.main+xml"
$templateContentType = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipEntryText {
    param(
        [string]$ZipPath,
        [string]$EntryName
    )

    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.GetEntry($EntryName)
        if ($null -eq $entry) {
            throw "Package entry not found: $EntryName"
        }

        $stream = $entry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
            try {
                return $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Set-ZipEntryText {
    param(
        [string]$ZipPath,
        [string]$EntryName,
        [string]$Content
    )

    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry($EntryName)
        if ($null -ne $entry) {
            $entry.Delete()
        }

        $entry = $zip.CreateEntry($EntryName)
        $stream = $entry.Open()
        try {
            $encoding = [System.Text.UTF8Encoding]::new($false)
            $writer = [System.IO.StreamWriter]::new($stream, $encoding)
            try {
                $writer.Write($Content)
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Convert-DocmPackageToDotmPackage {
    param(
        [string]$SourceDocmPath,
        [string]$TargetDotmPath
    )

    Copy-Item -LiteralPath $SourceDocmPath -Destination $TargetDotmPath -Force

    $contentTypes = Get-ZipEntryText -ZipPath $TargetDotmPath -EntryName "[Content_Types].xml"
    if (-not $contentTypes.Contains($documentContentType)) {
        throw "Source package does not contain the expected macro-enabled document content type."
    }

    $contentTypes = $contentTypes.Replace($documentContentType, $templateContentType)
    Set-ZipEntryText -ZipPath $TargetDotmPath -EntryName "[Content_Types].xml" -Content $contentTypes

    $appXml = Get-ZipEntryText -ZipPath $TargetDotmPath -EntryName "docProps/app.xml"
    $appXml = [regex]::Replace($appXml, "<Template>.*?</Template>", "<Template>CleanupSuite.dotm</Template>")
    Set-ZipEntryText -ZipPath $TargetDotmPath -EntryName "docProps/app.xml" -Content $appXml
}

New-Item -ItemType Directory -Force -Path $StartupDir | Out-Null
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

if (Test-Path -LiteralPath $startupDotmPath) {
    $backupPath = Join-Path $BackupDir ("CleanupSuite.startup-backup-$timestamp.dotm")
    Copy-Item -LiteralPath $startupDotmPath -Destination $backupPath -Force
    Write-Output "Backed up existing Startup template to $backupPath"
} else {
    Write-Output "No existing Startup template found at $startupDotmPath"
}

$tempDotmPath = Join-Path $env:TEMP ("CleanupSuite.startup-refresh-$timestamp.dotm")
Convert-DocmPackageToDotmPackage -SourceDocmPath $sourceDocmPath -TargetDotmPath $tempDotmPath

if (-not (Test-Path -LiteralPath $tempDotmPath)) {
    throw "Startup template package was not generated: $tempDotmPath"
}

Copy-Item -LiteralPath $tempDotmPath -Destination $startupDotmPath -Force
Write-Output "Refreshed Startup template at $startupDotmPath"

if (Test-Path -LiteralPath $tempDotmPath) {
    Remove-Item -LiteralPath $tempDotmPath -Force
}
