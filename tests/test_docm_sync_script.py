from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class DocmSyncScriptTests(unittest.TestCase):
    def test_sync_script_resolves_repo_root_dynamically(self):
        script = (ROOT / "scripts" / "sync_docm_code_only.ps1").read_text(encoding="utf-8")

        self.assertIn('$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path', script)
        self.assertNotIn('C:\\Users\\Chris\\PROGRAMMING\\VBA\\CleanupSuite for MS Word', script)

    def test_sync_script_includes_runtime_critical_components(self):
        script = (ROOT / "scripts" / "sync_docm_code_only.ps1").read_text(encoding="utf-8")

        self.assertIn('ComponentName = "modCleanupLauncher"', script)
        self.assertIn('ComponentName = "modAlphaSmokeRunner"', script)
        self.assertIn('ComponentName = "modMetaDataSuite"', script)
        self.assertIn('ComponentName = "frmPreviewActions"', script)
        self.assertIn('"cmdPreviousPage", "cmdNextPage", "lblButtonMeasure"', script)
        self.assertNotIn('"frmPreviewActions" = @("cmdPreviousPage", "cmdNextPage")', script)
        self.assertIn('ComponentName = "frmCleanupSuiteLauncher"', script)
        self.assertIn('ComponentName = "frmMetaDataSuite"', script)
        self.assertIn('ComponentName = "frmFinalReview"', script)
        self.assertIn('Ensure-FormComponent', script)
        self.assertIn('Ensure-StandardModuleComponent', script)
        self.assertIn('Sync-ControlCaptions', script)
        self.assertIn('if ($ControlName.StartsWith("lblRisk")) { return "Forms.CommandButton.1" }', script)
        self.assertIn('Remove-RecreatedFormControls', script)
        self.assertIn('"cmdHelpFinalReview"', script)
        self.assertIn('"cmdFinalReview"', script)
        self.assertIn('"lblDescFinalReview"', script)
        self.assertIn('"optHideLinks", "optRemoveLinks"', script)
        self.assertIn('"cmdRiskPlacementHyperlinkHide", "cmdRiskPlacementHyperlinkRemove"', script)
        self.assertIn('"cmdDocTrim" = "Trim End of Doc"', script)
        self.assertIn('"lblDescSoftReturn" = "Convert Shift+Enter line breaks to paragraph marks, or the reverse of that."', script)
        self.assertIn('"lblDescDuplicate" = "Find repeated or near-duplicate paragraphs before removing them."', script)
        self.assertIn('"lblDescHyperlink" = "Keep links but Hide styling`r`nKeep text but remove links"', script)
        self.assertNotIn('"lblDescHyperlink" = "Hide hyperlink styling while keeping links, or remove link targets while keeping visible text."', script)
        self.assertIn('"lblResetAllCalloutText" = "If highlights remain after an error, select Reset All."', script)
        self.assertIn('"lblRiskList" = "Structure`r`nchange"', script)
        self.assertIn('"lblRiskDuplicate" = "Inspect`r`nFirst"', script)
        self.assertIn('try { $control.TakeFocusOnClick = $false } catch {}', script)
        self.assertIn('Remove-ObsoleteFormControls', script)
        self.assertIn('"chkEnableAdvancedMetadataTools"', script)
        self.assertIn('"frmHyperlinkRemover" = @("chkRemoveFormat", "lblRiskPlacementHyperlinkWarning")', script)
        self.assertIn('"frmCapitalizationCleanup"', script)
        self.assertIn('"cmdEditExceptions"', script)

    def test_sync_script_creates_conformed_risk_controls_for_guided_forms(self):
        script = (ROOT / "scripts" / "sync_docm_code_only.ps1").read_text(encoding="utf-8")
        installer = (ROOT / "src" / "installer" / "installer.bas").read_text(encoding="utf-8")

        expected_controls = [
            "cmdRiskPlacementListAll", "cmdRiskPlacementListBullets", "cmdRiskPlacementListNumbering", "cmdRiskPlacementListIndent", "cmdRiskPlacementListCustom",
            "cmdRiskPlacementParagraphAll", "cmdRiskPlacementParagraphRemoveEmpty", "cmdRiskPlacementParagraphSpacing", "cmdRiskPlacementParagraphCustom",
            "cmdRiskPlacementDuplicatePreview", "cmdRiskPlacementDuplicateRemove", "cmdRiskPlacementDuplicateExact", "cmdRiskPlacementDuplicateNormalized", "cmdRiskPlacementDuplicateFuzzy",
            "cmdRiskPlacementFormattingChar", "cmdRiskPlacementFormattingPara", "cmdRiskPlacementFormattingQuick", "cmdRiskPlacementFormattingThorough", "cmdRiskPlacementFormattingStrip",
            "cmdRiskPlacementFontFace", "cmdRiskPlacementFontSize", "cmdRiskPlacementFontBold", "cmdRiskPlacementFontItalic", "cmdRiskPlacementFontColor",
            "cmdRiskPlacementStyleRemoveUnused", "cmdRiskPlacementStyleRemap",
            "cmdRiskPlacementTableRows", "cmdRiskPlacementTableCols", "cmdRiskPlacementTablePadding", "cmdRiskPlacementTableStripFormat", "cmdRiskPlacementTableBorders", "cmdRiskPlacementTableRemoveBorders", "cmdRiskPlacementTableConvertToText",
            "cmdRiskPlacementBreakSections", "cmdRiskPlacementBreakPages", "cmdRiskPlacementBreakConvert",
            "cmdRiskPlacementHeaderFooterStandardize", "cmdRiskPlacementHeaderFooterClear", "cmdRiskPlacementHeaderFooterBreakLinks",
            "cmdRiskPlacementFootnoteFootnotes", "cmdRiskPlacementFootnoteEndnotes", "cmdRiskPlacementFootnoteKeepText",
            "cmdRiskPlacementObjectPictures", "cmdRiskPlacementObjectTextBoxes", "cmdRiskPlacementObjectFrames", "cmdRiskPlacementObjectLines", "cmdRiskPlacementObjectControls", "cmdRiskPlacementObjectHiddenText", "cmdRiskPlacementObjectTables",
            "lblRiskPlacementDocumentTrimSummary",
            "cmdRiskPlacementSoftReturnToPara", "cmdRiskPlacementSoftReturnToSoft",
            "lblRiskPlacementFinalReviewSummary",
        ]

        for control_name in expected_controls:
            with self.subTest(control=control_name):
                self.assertIn(control_name, script)
                self.assertIn(control_name, installer)


if __name__ == "__main__":
    unittest.main()
