param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string[]]$TargetDocPaths = @()
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

function Get-ControlTypeByName {
    param([string]$ControlName)

    if ($ControlName.StartsWith("lblRisk")) { return "Forms.CommandButton.1" }
    if ($ControlName.StartsWith("opt")) { return "Forms.OptionButton.1" }
    if ($ControlName.StartsWith("chk")) { return "Forms.CheckBox.1" }
    if ($ControlName.StartsWith("cmd")) { return "Forms.CommandButton.1" }
    if ($ControlName.StartsWith("fra")) { return "Forms.Frame.1" }
    if ($ControlName.StartsWith("lst")) { return "Forms.ListBox.1" }
    if ($ControlName.StartsWith("cbo")) { return "Forms.ComboBox.1" }
    if ($ControlName.StartsWith("lbl")) { return "Forms.Label.1" }
    return "Forms.TextBox.1"
}

function Get-LauncherRiskControlNames {
    return @(
        "lblRiskUnicode", "lblRiskPunctuation", "lblRiskSpacing", "lblRiskCapitalization",
        "lblRiskList", "lblRiskParagraph", "lblRiskSoftReturn", "lblRiskDuplicate",
        "lblRiskTableClean", "lblRiskHeaderFooter", "lblRiskFootnote", "lblRiskObjectRemover",
        "lblRiskFontNorm", "lblRiskFormatStrip",
        "lblRiskMetadata", "lblRiskFinalReview", "lblRiskHyperlink",
        "lblRiskStyleClean", "lblRiskBreakNorm", "lblRiskDocTrim"
    )
}

function Get-ControlsForForm {
    param([string]$FormName)

    switch ($FormName) {
        "frmPreviewActions" {
            return @(
                "cmdPreviousChange", "cmdPreviousPage", "cmdNextPage", "cmdNextChange",
                "cmdReviewDetails", "lblReviewContext", "lblButtonMeasure"
            )
        }
        "frmCapitalizationCleanup" {
            return @(
                "cmdRiskPlacementCapitalizationRecommended",
                "cmdRiskPlacementCapitalizationCustom",
                "chkHeadingParentheses",
                "cmdEditExceptions"
            )
        }
        "frmCleanupSuiteLauncher" {
            return @(
                "lblCatShare",
                "lblCatAlpha"
            ) + (Get-LauncherRiskControlNames) + @(
                "cmdHelpFinalReview", "cmdFinalReview", "lblDescFinalReview",
                "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata",
                "lblAutoSaveWarning", "chkUpdateChecks", "cmdCheckUpdates",
                "lblResetAllArrow", "lblResetAllCalloutBox", "lblResetAllCalloutText"
            )
        }
        "frmFinalReview" {
            return @(
                "chkComments", "chkRevisions",
                "cmdAuthorCleanupMinimal", "cmdAuthorCleanupNormal", "cmdAuthorCleanupBroad",
                "lblRiskPlacementFinalReviewSummary",
                "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset",
                "fraScopeSelection", "optScopeDocument", "optScopeSelection"
            )
        }
        "frmHyperlinkRemover" {
            return @(
                "optHideLinks", "optRemoveLinks",
                "cmdRiskPlacementHyperlinkHide", "cmdRiskPlacementHyperlinkRemove"
            )
        }
        "frmListCleanup" {
            return @(
                "cmdRiskPlacementListAll", "cmdRiskPlacementListBullets", "cmdRiskPlacementListNumbering",
                "cmdRiskPlacementListIndent", "cmdRiskPlacementListCustom",
                "cmdRiskPlacementListNormalizeBullets", "cmdRiskPlacementListNormalizeNumbering",
                "cmdRiskPlacementListFixIndent", "cmdRiskPlacementListHyphenBullets"
            )
        }
        "frmParagraphCleanup" {
            return @(
                "cmdRiskPlacementParagraphAll", "cmdRiskPlacementParagraphRemoveEmpty",
                "cmdRiskPlacementParagraphSpacing", "cmdRiskPlacementParagraphCustom",
                "cmdRiskPlacementParagraphChkRemoveEmpty", "cmdRiskPlacementParagraphChkBreaks",
                "cmdRiskPlacementParagraphChkSpacing", "cmdRiskPlacementParagraphChkIndent"
            )
        }
        "frmDuplicateDetector" {
            return @(
                "cmdRiskPlacementDuplicateExact", "cmdRiskPlacementDuplicateNormalized",
                "cmdRiskPlacementDuplicateFuzzy", "chkIncludeEmptyParagraphs", "cmdRiskPlacementDuplicateEmpty",
                "cmdRiskPlacementDuplicateFuzzyLoose",
                "cmdRiskPlacementDuplicateFuzzyMedium", "cmdRiskPlacementDuplicateFuzzyStrict"
            )
        }
        "frmFontNormalizer" {
            return @(
                "cmdRiskPlacementFontFace", "cmdRiskPlacementFontSize", "cmdRiskPlacementFontBold",
                "cmdRiskPlacementFontItalic", "cmdRiskPlacementFontColor"
            )
        }
        "frmTableCleaner" {
            return @(
                "cmdRiskPlacementTableRows", "cmdRiskPlacementTableCols", "cmdRiskPlacementTablePadding",
                "cmdRiskPlacementTableStripFormat", "cmdRiskPlacementTableBorders",
                "cmdRiskPlacementTableRemoveBorders", "cmdRiskPlacementTableConvertToText"
            )
        }
        "frmBreakNormalizer" {
            return @(
                "cmdRiskPlacementBreakSections", "cmdRiskPlacementBreakPages", "cmdRiskPlacementBreakConvert",
                "cmdRiskPlacementBreakNextPage", "cmdRiskPlacementBreakContinuous",
                "cmdRiskPlacementBreakEvenPage", "cmdRiskPlacementBreakOddPage"
            )
        }
        "frmDocumentTrim" {
            return @(
                "lblRiskPlacementDocumentTrimSummary"
            )
        }
        "frmFormattingStripper" {
            return @(
                "cmdRiskPlacementFormattingChar", "cmdRiskPlacementFormattingPara",
                "cmdRiskPlacementFormattingQuick", "cmdRiskPlacementFormattingThorough",
                "cmdRiskPlacementFormattingStrip", "cmdRiskPlacementFormattingHighlight",
                "cmdRiskPlacementFormattingDropCaps"
            )
        }
        "frmSoftReturnConverter" {
            return @(
                "cmdRiskPlacementSoftReturnToPara", "cmdRiskPlacementSoftReturnToSoft"
            )
        }
        "frmStyleCleanup" {
            return @(
                "cmdRiskPlacementStyleRemoveUnused", "cmdRiskPlacementStyleRemap"
            )
        }
        "frmFootnoteRemover" {
            return @(
                "cmdRiskPlacementFootnoteFootnotes", "cmdRiskPlacementFootnoteEndnotes",
                "cmdRiskPlacementFootnoteKeepText"
            )
        }
        "frmHeaderFooterStandardizer" {
            return @(
                "cmdRiskPlacementHeaderFooterStandardize", "cmdRiskPlacementHeaderFooterClear",
                "cmdRiskPlacementHeaderFooterHeaders", "cmdRiskPlacementHeaderFooterFooters",
                "cmdRiskPlacementHeaderFooterFont", "cmdRiskPlacementHeaderFooterSpacing",
                "cmdRiskPlacementHeaderFooterBreakLinks", "cmdRiskPlacementHeaderFooterAlignment"
            )
        }
        "frmObjectRemover" {
            return @(
                "cmdRiskPlacementObjectPictures", "cmdRiskPlacementObjectTextBoxes",
                "cmdRiskPlacementObjectFrames", "cmdRiskPlacementObjectLines",
                "cmdRiskPlacementObjectControls", "cmdRiskPlacementObjectHiddenText",
                "cmdRiskPlacementObjectTables"
            )
        }
        "frmPunctuationCleanup" {
            return @(
                "cmdRiskPlacementPunctuationAll",
                "cmdRiskPlacementPunctuationQuotes",
                "cmdRiskPlacementPunctuationDashes",
                "cmdRiskPlacementPunctuationEllipses",
                "cmdRiskPlacementPunctuationCustom"
            )
        }
        "frmSpacingCleanup" {
            return @(
                "cmdRiskPlacementSpacingAll",
                "cmdRiskPlacementSpacingDouble",
                "cmdRiskPlacementSpacingTrim",
                "cmdRiskPlacementSpacingCustom"
            )
        }
        "frmUnicodeCleanup" {
            return @(
                "cmdRiskPlacementUnicodeAll",
                "cmdRiskPlacementUnicodeNBSP",
                "cmdRiskPlacementUnicodeZeroWidth",
                "cmdRiskPlacementUnicodeCustom"
            )
        }
        "frmMetaDataSuite" {
            return @(
                "lblTitle", "lblTargetTitle", "lblTargetPath", "lblOverallIndicator", "lblOverallStatus", "lblHint",
                "cmdRefreshCurrent", "cmdAuditAnother", "cmdOpenWordReport", "cmdSafeEditCurrent", "cmdGroupEditCurrent", "cmdClearSharingProperties", "cmdClose", "cmdExpandAll", "chkShowSuiteMetadata",
                "lblCoreHeader", "chkCoreDetails", "lblCoreIndicator", "lblCoreSummary", "txtCoreDetails",
                "lblAppHeader", "chkAppDetails", "lblAppIndicator", "lblAppSummary", "txtAppDetails",
                "lblCustomHeader", "chkCustomDetails", "lblCustomIndicator", "lblCustomSummary", "txtCustomDetails",
                "lblPackageHeader", "chkPackageDetails", "lblPackageIndicator", "lblPackageSummary", "txtPackageDetails",
                "lblMetadataHeader", "chkMetadataDetails", "lblMetadataIndicator", "lblMetadataSummary", "txtMetadataDetails", "cmdMetadataInvestigate",
                "lblCryptoHeader", "chkCryptoDetails", "lblCryptoIndicator", "lblCryptoSummary", "txtCryptoDetails", "cmdCryptoInvestigate"
            )
        }
        "frmSafeMetadataEditor" {
            return @(
                "lblEditorTitle", "lblDocumentPath", "lblSummaryHeader",
                "lblTitle", "txtTitle", "lblSubject", "txtSubject", "lblAuthor", "txtAuthor",
                "lblKeywords", "txtKeywords", "lblComments", "txtComments", "lblCategory", "txtCategory",
                "lblCompany", "txtCompany", "lblManager", "txtManager",
                "lblStatusHeader", "lblLastAuthor", "txtLastAuthor", "lblContentStatus", "txtContentStatus",
                "lblContentType", "txtContentType", "lblLanguage", "txtLanguage",
                "lblCustomHeader", "lblCustomHint", "lblCustomNameHeader", "lblCustomTypeHeader", "lblCustomValueHeader",
                "lstCustomProperties", "lblCustomName", "txtCustomName", "lblCustomType", "cboCustomType",
                "lblCustomValue", "txtCustomValue", "cmdCustomAddUpdate", "cmdCustomDelete",
                "lblAuditOnlyNote", "cmdSave", "cmdCancel"
            )
        }
        default {
            return @()
        }
    }
}

function Ensure-FormControls {
    param(
        [object]$Component,
        [string]$FormName
    )

    foreach ($controlName in (Get-ControlsForForm -FormName $FormName)) {
        $control = $null
        try {
            $control = $Component.Designer.Controls.Item($controlName)
        }
        catch {
        }

        if ($null -eq $control) {
            $null = $Component.Designer.Controls.Add((Get-ControlTypeByName -ControlName $controlName), $controlName, $true)
        }
    }
}

function Remove-RecreatedFormControls {
    param(
        [object]$Component,
        [string]$FormName
    )

    if ($FormName -ne "frmCleanupSuiteLauncher") {
        return
    }

    foreach ($controlName in (Get-LauncherRiskControlNames)) {
        try {
            $Component.Designer.Controls.Remove($controlName)
        }
        catch {
        }
    }
}

function Remove-ObsoleteFormControls {
    param(
        [object]$Component,
        [string]$FormName
    )

    $obsoleteControls = @{
        "frmCleanupSuiteLauncher" = @("chkEnableAdvancedMetadataTools", "chkShowCompletionReviewAfterApply")
        "frmHyperlinkRemover" = @("chkRemoveFormat", "lblRiskPlacementHyperlinkWarning")
    }

    if (-not $obsoleteControls.ContainsKey($FormName)) {
        return
    }

    foreach ($controlName in $obsoleteControls[$FormName]) {
        try {
            $Component.Designer.Controls.Remove($controlName)
        }
        catch {
        }
    }
}

function Ensure-FormComponent {
    param(
        [object]$Document,
        [string]$ComponentName
    )

    try {
        return $Document.VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        $component = $Document.VBProject.VBComponents.Add(3)
        $component.Name = $ComponentName

        Ensure-FormControls -Component $component -FormName $ComponentName

        return $component
    }
}

function Ensure-StandardModuleComponent {
    param(
        [object]$Document,
        [string]$ComponentName
    )

    try {
        return $Document.VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        $component = $Document.VBProject.VBComponents.Add(1)
        $component.Name = $ComponentName
        return $component
    }
}

function Sync-DocumentCode {
    param(
        [object]$Document,
        [object[]]$SourceItems
    )

    $vbProj = $Document.VBProject
    Remove-ObsoleteComponents -Document $Document

    foreach ($item in $SourceItems) {
        try {
            $component = $vbProj.VBComponents.Item($item.ComponentName)
        }
        catch {
            if ($item.ComponentName.StartsWith("frm")) {
                $component = Ensure-FormComponent -Document $Document -ComponentName $item.ComponentName
            }
            elseif ($item.ComponentName.StartsWith("mod")) {
                $component = Ensure-StandardModuleComponent -Document $Document -ComponentName $item.ComponentName
            }
            else {
                throw "Component $($item.ComponentName) not found in $($Document.FullName)"
            }
        }

        if ($item.ComponentName.StartsWith("frm")) {
            Remove-RecreatedFormControls -Component $component -FormName $item.ComponentName
            Remove-ObsoleteFormControls -Component $component -FormName $item.ComponentName
            Ensure-FormControls -Component $component -FormName $item.ComponentName
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

function Remove-ObsoleteComponents {
    param([object]$Document)

    $obsoleteComponentNames = @(
        "frmMetadataScrubber",
        "UserForm1",
        "UserForm2",
        "UserForm3"
    )

    foreach ($componentName in $obsoleteComponentNames) {
        try {
            $component = $Document.VBProject.VBComponents.Item($componentName)
            $Document.VBProject.VBComponents.Remove($component)
        }
        catch {
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
        "frmDuplicateDetector" = "Duplicate Paragraph Remover"
        "frmFontNormalizer" = "Font Normalizer"
        "frmTableCleaner" = "Table Cleaner"
        "frmBreakNormalizer" = "Break Normalizer"
        "frmDocumentTrim" = "Document Trim"
        "frmFormattingStripper" = "Formatting Stripper"
        "frmHyperlinkRemover" = "Hyperlink Cleaner"
        "frmSoftReturnConverter" = "Soft Return Converter"
        "frmMetaDataSuite" = "MetaDataSuite"
        "frmSafeMetadataEditor" = "Safe Metadata Editor"
        "frmFinalReview" = "Final Review"
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

function Sync-ControlCaptions {
    param([object]$Document)

    $captionByControl = @{
        "frmCleanupSuiteLauncher" = @{
            "chkAutoSave" = "Global default: auto-save before each tool"
            "chkUpdateChecks" = "Check periodically for updates"
            "cmdCheckUpdates" = "Check for Updates"
            "cmdDocTrim" = "Trim End of Doc"
            "lblDescSoftReturn" = "Convert Shift+Enter line breaks to paragraph marks, or the reverse of that."
            "lblDescDuplicate" = "Find repeated or near-duplicate paragraphs before removing them."
            "lblDescHyperlink" = "Keep links but Hide styling`r`nKeep text but remove links"
            "lblResetAllCalloutText" = "If highlights remain after an error, select Reset All."
            "lblAutoSaveWarning" = "It is dangerous to turn off auto save!"
            "lblRiskUnicode" = "Safe`r`ncleanup"
            "lblRiskPunctuation" = "Safe`r`ncleanup"
            "lblRiskSpacing" = "Safe`r`ncleanup"
            "lblRiskCapitalization" = "Text`r`ncaution"
            "lblRiskList" = "Structure`r`nchange"
            "lblRiskParagraph" = "Structure`r`nchange"
            "lblRiskSoftReturn" = "Structure`r`nchange"
            "lblRiskDuplicate" = "Inspect`r`nFirst"
            "lblRiskTableClean" = "Structure`r`nchange"
            "lblRiskHeaderFooter" = "Structure`r`nchange"
            "lblRiskFootnote" = "Removes`r`ncontent"
            "lblRiskObjectRemover" = "Removes`r`ncontent"
            "lblRiskFontNorm" = "Formatting`r`nchange"
            "lblRiskFormatStrip" = "Formatting`r`nchange"
            "lblRiskMetadata" = "Privacy`r`ncleanup"
            "lblRiskFinalReview" = "Finalizes`r`nreview"
            "lblRiskHyperlink" = "Removes`r`ncontent"
            "lblRiskStyleClean" = "Alpha`r`ntool"
            "lblRiskBreakNorm" = "Alpha`r`ntool"
            "lblRiskDocTrim" = "Safe`r`ncleanup"
        }
        "frmHyperlinkRemover" = @{
            "optHideLinks" = "Hide hyperlink styling, keep links"
            "optRemoveLinks" = "Remove hyperlinks, keep visible text"
        }
        "frmCapitalizationCleanup" = @{
            "chkHeadingParentheses" = "Also title-case words inside parentheses"
        }
    }

    foreach ($formName in $captionByControl.Keys) {
        try {
            $component = $Document.VBProject.VBComponents.Item($formName)
            foreach ($controlName in $captionByControl[$formName].Keys) {
                try {
                    $control = $component.Designer.Controls.Item($controlName)
                    $control.Caption = $captionByControl[$formName][$controlName]
                    if ($controlName.StartsWith("lblRisk")) {
                        try { $control.WordWrap = $true } catch {}
                        try { $control.TakeFocusOnClick = $false } catch {}
                    }
                }
                catch {
                }
            }
        }
        catch {
        }
    }
}

$sourceItems = @(
    @{ Path = (Join-Path $RepoRoot "src\forms\frmBreakNormalizer.bas"); ComponentName = "frmBreakNormalizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmCapitalizationCleanup.bas"); ComponentName = "frmCapitalizationCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmCleanupSuiteLauncher.bas"); ComponentName = "frmCleanupSuiteLauncher" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmDocumentTrim.bas"); ComponentName = "frmDocumentTrim" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmDuplicateDetector.bas"); ComponentName = "frmDuplicateDetector" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFontNormalizer.bas"); ComponentName = "frmFontNormalizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFootnoteRemover.bas"); ComponentName = "frmFootnoteRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFormattingStripper.bas"); ComponentName = "frmFormattingStripper" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmHeaderFooterStandardizer.bas"); ComponentName = "frmHeaderFooterStandardizer" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmHyperlinkRemover.bas"); ComponentName = "frmHyperlinkRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmListCleanup.bas"); ComponentName = "frmListCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmMetaDataSuite.bas"); ComponentName = "frmMetaDataSuite" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmSafeMetadataEditor.bas"); ComponentName = "frmSafeMetadataEditor" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmFinalReview.bas"); ComponentName = "frmFinalReview" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmObjectRemover.bas"); ComponentName = "frmObjectRemover" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmParagraphCleanup.bas"); ComponentName = "frmParagraphCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmPreviewActions.bas"); ComponentName = "frmPreviewActions" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmPunctuationCleanup.bas"); ComponentName = "frmPunctuationCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmSoftReturnConverter.bas"); ComponentName = "frmSoftReturnConverter" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmSpacingCleanup.bas"); ComponentName = "frmSpacingCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmStyleCleanup.bas"); ComponentName = "frmStyleCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmTableCleaner.bas"); ComponentName = "frmTableCleaner" },
    @{ Path = (Join-Path $RepoRoot "src\forms\frmUnicodeCleanup.bas"); ComponentName = "frmUnicodeCleanup" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modAlphaSmokeRunner.bas"); ComponentName = "modAlphaSmokeRunner" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modCleanupHelpers.bas"); ComponentName = "modCleanupHelpers" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modCleanupLauncher.bas"); ComponentName = "modCleanupLauncher" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modHybridBridge.bas"); ComponentName = "modHybridBridge" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modHybridJson.bas"); ComponentName = "modHybridJson" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modMetaDataSuite.bas"); ComponentName = "modMetaDataSuite" },
    @{ Path = (Join-Path $RepoRoot "src\modules\modUpdateChecker.bas"); ComponentName = "modUpdateChecker" },
    @{ Path = (Join-Path $RepoRoot "src\installer\installer.bas"); ComponentName = "modCleanupSuiteInstaller" }
)

$sourceItems = foreach ($item in $sourceItems) {
    [pscustomobject]@{
        Path = $item.Path
        ComponentName = $item.ComponentName
        CodeBody = Get-CodeBodyFromSource -Path $item.Path
    }
}

if ($TargetDocPaths.Count -gt 0) {
    $targetDocs = $TargetDocPaths
} else {
    $targetDocs = @(
        (Join-Path $RepoRoot "CleanupSuite.docm"),
        (Join-Path $RepoRoot "Practice - Try CleanupSuite Here\CleanupSuite.docm")
    )
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($targetDoc in $targetDocs) {
        $targetExtension = [IO.Path]::GetExtension($targetDoc)
        if ([string]::IsNullOrWhiteSpace($targetExtension)) {
            $targetExtension = ".docm"
        }
        $tempPath = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $targetExtension)
        Copy-Item -LiteralPath $targetDoc -Destination $tempPath -Force

        $document = $null
        try {
            $document = $word.Documents.Open($tempPath, $false, $false)
            Sync-DocumentCode -Document $document -SourceItems $sourceItems
            Sync-FormCaptions -Document $document
            Sync-ControlCaptions -Document $document
            try {
                $document.ActiveWindow.View.ReadingLayout = $false
                $document.ActiveWindow.View.Type = 3
            } catch {}
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
