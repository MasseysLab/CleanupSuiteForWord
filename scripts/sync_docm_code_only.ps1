param(
    [string]$RepoRoot = "C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word"
)

$ErrorActionPreference = "Stop"

function Get-CodeBodyFromSource {
    param([string]$Path)
    $lines = Get-Content -Path $Path
    $filtered = foreach ($line in $lines) {
        if ($line -match '^Attribute VB_Name = ') { continue }
        if ($line -match '^Attribute VB_(Description|Exposed|PredeclaredId|GlobalName|Creatable|TemplateDerived|Customizable) = ') { continue }
        $line
    }
    return ($filtered -join "`r`n")
}

function Sync-DocumentCode {
    param(
        [object]$Document,
        [object[]]$SourceItems
    )

    $vbProj = $Document.VBProject
    foreach ($item in $SourceItems) {
        $component = $vbProj.VBComponents.Item($item.ComponentName)
        if (-not $component) {
            throw "Component $($item.ComponentName) not found in $($Document.FullName)"
        }

        $codeModule = $component.CodeModule
        $lineCount = $codeModule.CountOfLines
        if ($lineCount -gt 0) {
            $codeModule.DeleteLines(1, $lineCount)
        }

        if ($item.CodeBody.Trim().Length -gt 0) {
            $null = $codeModule.AddFromString($item.CodeBody)
        }
    }
}

function Sync-FormCaptions {
    param([object]$Document)

    $captionByForm = [ordered]@{
        "frmCapitalizationCleanup" = "Capitalization Fixer"
        "frmPunctuationCleanup" = "Punctuation Normalizer"
        "frmUnicodeCleanup" = "Invisible Unicode Cleaner"
        "frmSpacingCleanup" = "Spacing Fixer"
        "frmListCleanup" = "List Normalizer"
        "frmParagraphCleanup" = "Paragraph Structure Fixer"
        "frmDuplicateDetector" = "Duplicate Paragraph Detector"
        "frmFontNormalizer" = "Font Normalizer"
        "frmTableCleaner" = "Table Cleaner"
        "frmBreakNormalizer" = "Break Normalizer"
        "frmDocumentTrim" = "Document Trim"
        "frmFormattingStripper" = "Formatting Stripper"
        "frmHyperlinkRemover" = "Hyperlink Remover"
        "frmSoftReturnConverter" = "Soft Return Converter"
        "frmMetadataScrubber" = "Metadata Scrubber"
        "frmStyleCleanup" = "Style Cleanup"
        "frmFootnoteRemover" = "Footnote / Endnote Remover"
        "frmHeaderFooterStandardizer" = "Header / Footer Standardizer"
        "frmObjectRemover" = "Object Remover"
    }

    foreach ($formName in $captionByForm.Keys) {
        $component = $Document.VBProject.VBComponents.Item($formName)
        if ($component -and $component.Properties) {
            try {
                $component.Properties("Caption").Value = $captionByForm[$formName]
            }
            catch {
            }
        }
        try {
            $component.Designer.Caption = $captionByForm[$formName]
        }
        catch {
        }
    }
}

$sourceItems = @(
    @{ Path = (Join-Path $RepoRoot "src\forms\frmBreakNormalizer.bas"); ComponentName = "frmBreakNormalizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmCapitalizationCleanup.bas"); ComponentName = "frmCapitalizationCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmDocumentTrim.bas"); ComponentName = "frmDocumentTrim" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmDuplicateDetector.bas"); ComponentName = "frmDuplicateDetector" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFontNormalizer.bas"); ComponentName = "frmFontNormalizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFootnoteRemover.bas"); ComponentName = "frmFootnoteRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFormattingStripper.bas"); ComponentName = "frmFormattingStripper" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmHeaderFooterStandardizer.bas"); ComponentName = "frmHeaderFooterStandardizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmHyperlinkRemover.bas"); ComponentName = "frmHyperlinkRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmListCleanup.bas"); ComponentName = "frmListCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmMetadataScrubber.bas"); ComponentName = "frmMetadataScrubber" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmObjectRemover.bas"); ComponentName = "frmObjectRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmParagraphCleanup.bas"); ComponentName = "frmParagraphCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmPunctuationCleanup.bas"); ComponentName = "frmPunctuationCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmSoftReturnConverter.bas"); ComponentName = "frmSoftReturnConverter" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmSpacingCleanup.bas"); ComponentName = "frmSpacingCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmStyleCleanup.bas"); ComponentName = "frmStyleCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmTableCleaner.bas"); ComponentName = "frmTableCleaner" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmUnicodeCleanup.bas"); ComponentName = "frmUnicodeCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modCleanupHelpers.bas"); ComponentName = "modCleanupHelpers" },
    @{ Path = (Join-Path $RepoRoot "src\installer\installer.bas"); ComponentName = "modCleanupSuiteInstaller" }
)

$sourceItems = foreach ($item in $sourceItems) {
    [pscustomobject]@{
        Path = $item.Path
        ComponentName = $item.ComponentName
        CodeBody = Get-CodeBodyFromSource -Path $item.Path
    }
}

$targetDocs = @(
    (Join-Path $RepoRoot "CleanupSuite_Workspace.docm"),
    (Join-Path $RepoRoot "Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm")
)

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($targetDoc in $targetDocs) {
        $tempPath = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + ".docm")
        Copy-Item -LiteralPath $targetDoc -Destination $tempPath -Force

        $document = $null
        try {
            $document = $word.Documents.Open($tempPath, $false, $false)
            Sync-DocumentCode -Document $document -SourceItems $sourceItems
            Sync-FormCaptions -Document $document
            $document.Save()
            $document.Close()
            $document = $null
            Copy-Item -LiteralPath $tempPath -Destination $targetDoc -Force
            Write-Output "Synced $targetDoc"
        }
        finally {
            if ($document -ne $null) {
                $document.Close([ref]$false)
            }
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force
            }
        }
    }
}
finally {
    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
