param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [switch]$BuildFirst
)

$ErrorActionPreference = "Stop"

function Invoke-RepoPowerShell {
    param([string]$RelativeScript)
    & powershell -ExecutionPolicy Bypass -File (Join-Path $RepoRoot $RelativeScript)
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $RelativeScript"
    }
}

function Get-VbComponent {
    param(
        [object]$Document,
        [string]$ComponentName
    )

    try {
        return $Document.VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        return $null
    }
}

function Get-CodeText {
    param([object]$Component)

    if ($null -eq $Component) {
        return ""
    }

    $module = $Component.CodeModule
    return $module.Lines(1, $module.CountOfLines)
}

function Test-ContainsAll {
    param(
        [string]$Text,
        [string[]]$Snippets
    )

    foreach ($snippet in $Snippets) {
        if ($Text -notlike "*$snippet*") {
            return $false
        }
    }

    return $true
}

function Test-ControlNames {
    param(
        [object]$Component,
        [string[]]$ControlNames
    )

    foreach ($controlName in $ControlNames) {
        try {
            $null = $Component.Designer.Controls.Item($controlName)
        }
        catch {
            return $false
        }
    }

    return $true
}

function New-ResultLine {
    param(
        [string]$Status,
        [string]$Label,
        [string]$Detail
    )

    return "$Status|$Label|$Detail"
}

if ($BuildFirst) {
    Invoke-RepoPowerShell "build.ps1"
}

Invoke-RepoPowerShell "scripts\sync_docm_code_only.ps1"

$suiteDocPath = Join-Path $RepoRoot "Practice - Try CleanupSuite Here\CleanupSuite.docm"
if (-not (Test-Path -LiteralPath $suiteDocPath)) {
    throw "Practice CleanupSuite docm not found: $suiteDocPath"
}

$reportDir = Join-Path $RepoRoot "docs\test-results"
New-Item -ItemType Directory -Force -Path $reportDir | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportPath = Join-Path $reportDir "alpha_smoke_report_$timestamp.txt"

# The repo also carries a VBA-side runtime hook named RunAlphaSmokeChecklist.
# The default runner uses safer embedded-docm inspection through Word COM.

$word = $null
$suiteDoc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 1 } catch {}

    $suiteDoc = $word.Documents.Open($suiteDocPath, $false, $true, $false)
    $lines = @()

    $launcher = Get-VbComponent -Document $suiteDoc -ComponentName "frmCleanupSuiteLauncher"
    $launcherCode = Get-CodeText -Component $launcher
    if ($launcher -and (Test-ControlNames $launcher @("cmdDocTrim", "cmdUnicode", "chkReturnToMainAfterApply")) -and (Test-ContainsAll $launcherCode @("Private Sub OpenCleanupTool(ByVal formName As String)", 'Private Sub cmdDocTrim_Click(): OpenCleanupTool "frmDocumentTrim": End Sub'))) {
        $lines += New-ResultLine "PASS" "Launcher" "The embedded practice docm contains frmCleanupSuiteLauncher with the expected launch controls and open-tool wiring."
    } else {
        $lines += New-ResultLine "FAIL" "Launcher" "frmCleanupSuiteLauncher is missing required controls or open-tool wiring in the practice docm."
    }

    $documentTrim = Get-VbComponent -Document $suiteDoc -ComponentName "frmDocumentTrim"
    $documentTrimCode = Get-CodeText -Component $documentTrim
    if ($documentTrim -and (Test-ControlNames $documentTrim @("cmdPreview", "cmdRun", "cmdReset", "optScopeDocument", "optScopeSelection")) -and (Test-ContainsAll $documentTrimCode @("Entire document (always)", "Public Sub PreviewFromPanel()", "Public Sub RunAfterPreview()"))) {
        $lines += New-ResultLine "PASS" "Document Trim" "The simplest guided form keeps the preview-only visible action flow in the embedded practice docm."
    } else {
        $lines += New-ResultLine "FAIL" "Document Trim" "Document Trim is missing expected guided-form controls or preview wiring in the practice docm."
    }

    $unicode = Get-VbComponent -Document $suiteDoc -ComponentName "frmUnicodeCleanup"
    $unicodeCode = Get-CodeText -Component $unicode
    if ($unicode -and (Test-ControlNames $unicode @("optAll", "optCustom", "chkNBSP", "chkZWSP")) -and (Test-ContainsAll $unicodeCode @('optCustom.Caption = "Custom"', "Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutUnicodeForm: End Sub", "Private Sub chkNBSP_Click(): LayoutUnicodeForm: End Sub"))) {
        $lines += New-ResultLine "PASS" "Invisible Unicode Cleaner" "The Custom-mode reference form keeps its option wiring and relayout hooks in the embedded practice docm."
    } else {
        $lines += New-ResultLine "FAIL" "Invisible Unicode Cleaner" "Invisible Unicode Cleaner is missing expected Custom-mode controls or relayout hooks in the practice docm."
    }

    $capitalization = Get-VbComponent -Document $suiteDoc -ComponentName "frmCapitalizationCleanup"
    $capitalizationCode = Get-CodeText -Component $capitalization
    if ($capitalization -and (Test-ControlNames $capitalization @("optAll", "optCustom", "chkHeadingParentheses", "chkSmartSentences", "cmdEditExceptions")) -and (Test-ContainsAll $capitalizationCode @('optAll.Caption = "Recommended repair"', 'optCustom.Caption = "Custom repair"', 'chkHeadingParentheses.Caption = "Also title-case words inside parentheses"', 'cmdEditExceptions.Caption = "Edit custom exceptions"', 'lblIntro.Caption = "Use the recommended repair, or customize its protections."'))) {
        $lines += New-ResultLine "PASS" "Capitalization Fixer" "The reference guided form keeps the smart-repair labels, custom-exceptions action, and advanced-protection controls in the embedded practice docm."
    } else {
        $lines += New-ResultLine "FAIL" "Capitalization Fixer" "Capitalization Fixer is missing expected smart-repair labels, custom-exceptions action, or advanced-protection controls in the practice docm."
    }

    $headerFooter = Get-VbComponent -Document $suiteDoc -ComponentName "frmHeaderFooterStandardizer"
    $headerFooterCode = Get-CodeText -Component $headerFooter
    if ($headerFooter -and (Test-ControlNames $headerFooter @("optStandardize", "optClearAll", "chkBreakLinks", "chkAlignment", "optAlignLeft", "optAlignCenter", "optAlignRight")) -and (Test-ContainsAll $headerFooterCode @('chkBreakLinks.Caption = "Unlink sections (set each independently)"', 'chkAlignment.Caption = "Set alignment"', "LayoutCleanupToolForm Me"))) {
        $lines += New-ResultLine "PASS" "Header / Footer Standardizer" "The special-case layout form keeps its unlink and alignment controls in the embedded practice docm."
    } else {
        $lines += New-ResultLine "FAIL" "Header / Footer Standardizer" "Header / Footer Standardizer is missing expected special-case layout controls in the practice docm."
    }

    $formattingCleaner = Get-VbComponent -Document $suiteDoc -ComponentName "frmFormattingStripper"
    $formattingCleanerCode = Get-CodeText -Component $formattingCleaner
    if ($formattingCleaner -and (Test-ControlNames $formattingCleaner @("chkResetChar", "chkResetPara", "cmdPreview", "cmdRun")) -and (Test-ContainsAll $formattingCleanerCode @("p.Range.Font.Reset", "p.Range.ParagraphFormat.Reset")) -and $formattingCleanerCode -notlike "*ClearCharacterDirectFormatting*" -and $formattingCleanerCode -notlike "*ClearParagraphDirectFormatting*") {
        $lines += New-ResultLine "PASS" "Formatting Cleaner" "The embedded practice docm uses the supported character and paragraph formatting reset APIs."
    } else {
        $lines += New-ResultLine "FAIL" "Formatting Cleaner" "Formatting Cleaner is missing required controls or still contains an unsupported direct-formatting API."
    }

    $objectRemover = Get-VbComponent -Document $suiteDoc -ComponentName "frmObjectRemover"
    $objectRemoverCode = Get-CodeText -Component $objectRemover
    $helpers = Get-VbComponent -Document $suiteDoc -ComponentName "modCleanupHelpers"
    $helpersCode = Get-CodeText -Component $helpers
    if ($objectRemover -and $headerFooter -and $helpers -and (Test-ControlNames $objectRemover @("chkTables", "cmdRiskPlacementObjectTables", "cmdRiskPlacementObjectPictures", "cmdPreview", "cmdRun")) -and (Test-ControlNames $headerFooter @("cmdRiskPlacementHeaderFooterClear", "cmdRiskPlacementHeaderFooterBreakLinks")) -and (Test-ContainsAll $objectRemoverCode @("Private Sub cmdRiskPlacementObjectTables_Click()", "ShowToolRiskChoiceExplanation")) -and (Test-ContainsAll $helpersCode @("Private Sub ApplyGuidedRiskPlacementLayout", "LayoutObjectRiskChoiceChips toolForm", "LayoutHeaderFooterRiskChoiceChips toolForm"))) {
        $lines += New-ResultLine "PASS" "Risk Controls" "High-impact guided forms carry conformed risk controls through the embedded practice docm."
    } else {
        $lines += New-ResultLine "FAIL" "Risk Controls" "High-impact guided forms are missing conformed risk controls in the embedded practice docm."
    }

    $previewForm = Get-VbComponent -Document $suiteDoc -ComponentName "frmPreviewActions"
    $previewFormCode = Get-CodeText -Component $previewForm
    if ($previewForm -and (Test-ControlNames $previewForm @("cmdApply", "cmdPreview", "cmdClear", "lblSummary", "lblHint")) -and (Test-ContainsAll $previewFormCode @('cmdApply.Caption = "Apply"', 'cmdClear.Caption = "Reconfigure"', 'cmdPreview.Caption = "Preview is ON"')) -and (Test-ContainsAll $documentTrimCode @('ShowPreviewActionsSummary Me, "Document Trim"', "Public Sub PreviewFromPanel()", "Public Sub RunAfterPreview()")) -and (Test-ContainsAll $helpersCode @("Public Sub ShowPreviewActionsSummary(ByVal sourceForm As Object, ByVal toolName As String, ByVal rows As Collection)", "Set gPreviewActionPanel = New frmPreviewActions", "gPreviewActionPanel.ConfigureSummary sourceForm, toolName, rows"))) {
        $lines += New-ResultLine "PASS" "Preview Flow" "The embedded practice docm keeps the Preview contract between Document Trim, frmPreviewActions, and modCleanupHelpers."
    } else {
        $lines += New-ResultLine "FAIL" "Preview Flow" "The embedded practice docm is missing part of the Preview contract between the tool form, preview panel, and helpers."
    }

    if ($previewForm -and (Test-ContainsAll $previewFormCode @("Public Sub ConfigureSummary(sourceForm As Object, toolName As String, rows As Collection)", "RenderSummaryTable rows", 'RenderSummaryCell "lblSummaryHeaderItem", "Item"', 'RenderSummaryCell "lblSummaryHeaderCount", "#"', "tableLeft = M + ((CONTENT_W - TABLE_W) / 2)", "SUMMARY_BOTTOM_GAP", "If Me.InsideHeight < desiredInsideH Then")) -and (Test-ContainsAll $helpersCode @("Public Function NewPreviewSummaryRows() As Collection", "Public Sub AddPreviewSummaryRow", '" (unhighlighted)"'))) {
        $lines += New-ResultLine "PASS" "Preview Summary Table" "Preview panel code renders centered Item/# rows, unhighlighted item suffixes, and inside-height clipping protection."
    } else {
        $lines += New-ResultLine "FAIL" "Preview Summary Table" "Preview panel code is missing the summary table contract."
    }

    $lines += New-ResultLine "WARN" "Visual review" "This runner validates the embedded practice docm wiring through Word COM. Do one brief live visual review for duplicate titles, panel cleanliness, and wasted vertical space before a release call."

    # Example lines: PASS|Launcher|..., PASS|Formatting Cleaner|..., PASS|Risk Controls|..., PASS|Preview Summary Table|..., and WARN|Visual review|...
    Set-Content -LiteralPath $reportPath -Value $lines -Encoding utf8
    foreach ($line in $lines) {
        Write-Host $line
    }
    Write-Host "Report: $reportPath"

    if ($lines | Where-Object { $_ -like "FAIL|*" }) {
        throw "Alpha smoke runner reported one or more FAIL lines."
    }
}
finally {
    if ($suiteDoc -ne $null) {
        try {
            $suiteDoc.ActiveWindow.View.ReadingLayout = $false
            $suiteDoc.ActiveWindow.View.Type = 3
        } catch {}
        $suiteDoc.Close([ref]$false)
    }
    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
