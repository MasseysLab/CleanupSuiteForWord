from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class MetaDataSuiteClearSharingPropertiesTests(unittest.TestCase):
    def test_dashboard_exposes_clear_sharing_properties_action(self):
        form = read("src/forms/frmMetaDataSuite.bas")
        module = read("src/modules/modMetaDataSuite.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("Private Sub cmdClearSharingProperties_Click()", form)
        self.assertIn("ClearMetadataDashboardSharingProperties Me", form)
        self.assertIn('frm.cmdClearSharingProperties.Caption = "Clear Sharing Properties"', module)
        self.assertIn("StyleMetadataDashboardButton frm.cmdClearSharingProperties, False", module)
        self.assertIn('frm.cmdClearSharingProperties.Move 12, 86, 180, 24', module)
        self.assertIn('"cmdClearSharingProperties"', installer)

    def test_clear_sharing_properties_previews_then_uses_word_inspector_actions(self):
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn("Public Sub ClearMetadataDashboardSharingProperties(ByVal frm As Object)", module)
        self.assertIn("previewText = BuildSharingPropertiesPreview(doc)", module)
        self.assertIn('MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton2, "Clear Sharing Properties")', module)
        self.assertIn("ApplyClearSharingProperties doc", module)
        self.assertIn("doc.RemoveDocumentInformation wdRDIDocumentProperties", module)
        self.assertIn("doc.RemoveDocumentInformation wdRDIRemovePersonalInformation", module)
        self.assertIn("RefreshMetadataDashboardQuickScan frm", module)

    def test_clear_sharing_properties_preserves_cleanupsuite_managed_custom_properties(self):
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn("Private Function IsCleanupSuiteManagedProperty(ByVal propertyName As String) As Boolean", module)
        for managed_name in [
            "CleanupSuiteAutoSave",
            "CleanupSuiteReturnToMainAfterApply",
            "CleanupSuiteReturnToMainAfterClose",
            "CleanupSuiteInProgress",
            "CleanupSuiteCapitalizationExceptions",
        ]:
            with self.subTest(property=managed_name):
                self.assertIn(f'Case "{managed_name}"', module)
        self.assertNotIn('Case "CleanupSuiteShowCompletionReviewAfterApply"', module)
        self.assertIn("Set protectedProps = SnapshotCleanupSuiteManagedProperties(doc)", module)
        self.assertIn("RestoreCleanupSuiteManagedProperties doc, protectedProps", module)

    def test_final_review_does_not_own_document_property_clearing(self):
        final_review = read("src/forms/frmFinalReview.bas")
        self.assertNotIn("RemoveDocumentInformation wdRDIDocumentProperties", final_review)
        self.assertNotIn("ClearMetadataDashboardSharingProperties", final_review)

    def test_legacy_metadata_scrubber_is_not_shipped_after_absorption(self):
        manifest = read("src/manifest.txt")
        installer = read("src/installer/installer.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        self.assertNotIn("frmMetadataScrubber_Code", manifest)
        self.assertNotIn('CreateOrReplaceForm vbProj, "frmMetadataScrubber"', installer)
        self.assertNotIn('AttemptCreateControls vbProj, "frmMetadataScrubber"', installer)
        self.assertNotIn('Case "frmMetadataScrubber"', installer)
        self.assertNotIn("frmMetadataScrubber", helpers)
        self.assertNotIn('ComponentName = "frmMetadataScrubber"', sync_script)
        self.assertIn('Remove-ObsoleteComponents -Document $Document', sync_script)

    def test_sync_removes_orphan_default_userforms(self):
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for component_name in ["UserForm1", "UserForm2", "UserForm3"]:
            with self.subTest(component=component_name):
                self.assertIn(f'"{component_name}"', sync_script)


if __name__ == "__main__":
    unittest.main()
