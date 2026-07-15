from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class MetaDataSuiteJump3Tests(unittest.TestCase):
    def test_launcher_uses_metadatasuite_name_without_advanced_setting(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for snippet in [
            'cmdMetadata.Caption = "MetaDataSuite"',
            'lblDescMetadata.Caption = "Advanced metadata inspection, editing, and privacy management."',
            "cmdMetadata.Visible = True",
            "lblRiskMetadata.Visible = True",
            'OpenCleanupTool "frmMetaDataSuite"',
            'ShowCleanupToolHelp "MetaDataSuite"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, launcher)

        self.assertIn('Case "MetaDataSuite": CleanupHelpTitle = "MetaDataSuite"', helpers)

        for source_name, source in [
            ("launcher", launcher),
            ("helpers", helpers),
            ("installer", installer),
        ]:
            for snippet in [
                "chkEnableAdvancedMetadataTools",
                "GetEnableAdvancedMetadataToolsSetting",
                "SetEnableAdvancedMetadataToolsSetting",
                "ConfirmEnableAdvancedMetadataTools",
                "CleanupSuiteEnableAdvancedMetadataTools",
                "CleanupSuiteAdvancedMetadataWarningShown",
            ]:
                with self.subTest(source=source_name, snippet=snippet):
                    self.assertNotIn(snippet, source)

        for snippet in [
            '"cmdHelpMetadata"',
            '"cmdMetadata"',
            '"lblDescMetadata"',
            'Remove-ObsoleteFormControls',
            '"chkEnableAdvancedMetadataTools"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, sync_script)

    def test_metadatasuite_form_and_module_exist(self):
        form_source = read("src/forms/frmMetaDataSuite.bas")
        module_source = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            "InitializeMetadataDashboard Me",
            "RefreshMetadataDashboardQuickScan Me",
            "OpenMetadataDashboardWordReport Me",
            "SafeEditMetadataDashboardTarget Me",
            "GroupEditMetadataDashboardTarget Me",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form_source)

        for snippet in [
            'Set mDashboardForm = New frmMetaDataSuite',
            "Public Sub ShowMetaDataSuiteDashboard()",
            'frm.lblTitle.Caption = "MetaDataSuite"',
            'frm.cmdRefreshCurrent.Caption = "Refresh Quick Checks"',
            'frm.chkCoreDetails.Caption = "Show details"',
            'frm.lblCoreHeader.Caption = "Core Properties"',
            'frm.cmdMetadataInvestigate.Caption = "Investigate"',
            "HideLegacyAuthorClearPanel frm",
            'txt.MultiLine = True',
            'txt.EnterKeyBehavior = True',
            'txt.Locked = True',
            'lineText = lineText & item.DisplayName & vbCrLf & _',
            'If NodeHasElementChildren(node) Then GoTo NextDocPropNode',
            "Public Sub RunFinalReviewAuthorCleanup(ByVal removalLevel As String)",
            "Public Function PreviewFinalReviewAuthorCleanup(ByVal removalLevel As String) As String",
            "AUTHOR_CLEAR_MINIMAL",
            "AUTHOR_CLEAR_NORMAL",
            "AUTHOR_CLEAR_BROAD",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, module_source)

    def test_metadatasuite_uses_cleanupsuite_theme_colors(self):
        module_source = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            "frm.BackColor = RGB(248, 249, 251)",
            "frm.lblTitle.ForeColor = RGB(32, 37, 45)",
            "StyleMetadataDashboardButton frm.cmdRefreshCurrent, False",
            "StyleMetadataDashboardButton frm.cmdExpandAll, True",
            "StyleMetadataDashboardSectionHeader frm.lblCoreHeader",
            "StyleMetadataDashboardSummary frm.lblCoreSummary",
            "StyleMetadataDashboardDetails frm.txtCoreDetails",
            "StyleMetadataDashboardCheckBox frm.chkShowSuiteMetadata",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, module_source)

    def test_final_review_hosts_author_cleanup_and_undoability_text(self):
        final_review = read("src/forms/frmFinalReview.bas")

        for snippet in [
            "cmdAuthorCleanupMinimal",
            "cmdAuthorCleanupNormal",
            "cmdAuthorCleanupBroad",
            'RunFinalReviewAuthorCleanup "minimal"',
            'RunFinalReviewAuthorCleanup "normal"',
            'RunFinalReviewAuthorCleanup "broad"',
            "Undoability: partially undoable or not reliably undoable.",
            "Undoability: effectively final; treat as permanent once saved.",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, final_review)

    def test_installer_manifest_and_sync_script_include_metadatasuite(self):
        installer = read("src/installer/installer.bas")
        manifest = read("src/manifest.txt")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for snippet in [
            'CreateOrReplaceModule vbProj, "modMetaDataSuite", modMetaDataSuite_Code, installLog',
            'CreateOrReplaceForm vbProj, "frmMetaDataSuite", frmMetaDataSuite_Code, installLog',
            'AttemptCreateControls vbProj, "frmMetaDataSuite", uiFailures',
            '"frmMetaDataSuite"',
            '"modMetaDataSuite"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, installer)

        for snippet in [
            "BUILDER  modules/modMetaDataSuite.bas  modMetaDataSuite_Code",
            "BUILDER  forms/frmMetaDataSuite.bas  frmMetaDataSuite_Code",
            "BUILDER  forms/frmSafeMetadataEditor.bas  frmSafeMetadataEditor_Code",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, manifest)

        for snippet in [
            'ComponentName = "frmMetaDataSuite"',
            'ComponentName = "frmSafeMetadataEditor"',
            'ComponentName = "modMetaDataSuite"',
            '"cmdMetadata"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, sync_script)


if __name__ == "__main__":
    unittest.main()
