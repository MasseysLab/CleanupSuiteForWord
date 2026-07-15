from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class MetaDataSuiteGroupEditSafetyTests(unittest.TestCase):
    def test_group_edit_current_open_document_is_view_only(self):
        source = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            "Private Function IsOpenDocumentPath(ByVal filePath As String) As Boolean",
            "RunPackageMetadataGroupViewOnlyWorkflow tempDir, filePath,",
            '"Group Edit Current is view-only for open documents."',
            "If IsOpenDocumentPath(filePath) Or IsRunningMacroHostPath(filePath) Then",
            '"The extracted package is open for inspection only."',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, source)

        self.assertIn("RunPackageMetadataGroupWorkflow tempDir, filePath", source)

    def test_cleanup_suite_self_metadata_is_hidden_by_default_with_toggle(self):
        form = read("src/forms/frmMetaDataSuite.bas")
        module = read("src/modules/modMetaDataSuite.bas")
        installer = read("src/installer/installer.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for snippet in [
            "Private mDashboardShowSuiteMetadata As Boolean",
            'frm.chkShowSuiteMetadata.Caption = "Show CleanupSuite metadata"',
            "frm.chkShowSuiteMetadata.Value = False",
            "Private Function FilterDashboardMetadataItems(",
            "ShouldHideSuiteMetadataItem",
            "NormalizeAuditTempPathText",
            "BuildPackageRelativePath",
            "MetadataToolkitRole",
            "word\\vbadata.xml",
            "cleanupsuitereturntomainafterapply",
            "cleanupsuitereturntomainafterclose",
            "If Not mDashboardShowSuiteMetadata Then",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, module)

        self.assertNotIn("CleanupSuiteShowCompletionReviewAfterApply", module)

        self.assertIn("Private Sub chkShowSuiteMetadata_Click()", form)
        self.assertIn("SetMetadataDashboardShowSuiteMetadata Me, chkShowSuiteMetadata.Value", form)
        self.assertIn('"chkShowSuiteMetadata"', installer)
        self.assertIn('Case "frmMetaDataSuite.chkShowSuiteMetadata": ControlCaptionText = "Show CleanupSuite metadata"', installer)
        self.assertIn('"chkShowSuiteMetadata"', sync_script)


if __name__ == "__main__":
    unittest.main()
