from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class SafeMetadataEditorTests(unittest.TestCase):
    def test_editor_groups_complete_safe_property_whitelist(self):
        module = read("src/modules/modMetaDataSuite.bas")
        form = read("src/forms/frmSafeMetadataEditor.bas")

        for property_name in [
            "Title", "Subject", "Author", "Keywords", "Comments",
            "Category", "Company", "Manager", "Last Author",
            "Content Status", "Content Type", "Language",
        ]:
            with self.subTest(property_name=property_name):
                self.assertIn(f'"{property_name}"', module)

        for section in [
            '"Document summary"',
            '"Attribution and classification"',
            '"Custom document properties"',
            '"Audit-only:',
        ]:
            with self.subTest(section=section):
                self.assertIn(section, module)

        self.assertIn("InitializeSafeMetadataEditor Me", form)
        self.assertIn("SaveSafeMetadataEditor Me", form)
        self.assertIn("AddOrUpdateSafeCustomProperty Me", form)
        self.assertIn("DeleteSelectedSafeCustomProperty Me", form)

    def test_editor_fits_the_launcher_1024_by_768_screen_envelope(self):
        module = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            "Const SAFE_EDITOR_W As Single = 760",
            "Const SAFE_EDITOR_H As Single = 510",
            "These are VBA form points, not screen pixels.",
            ".Move 24, 356, 530, 56",
            'ConfigureSafeEditorButton frm.cmdSave, "Save safe metadata", 456, 448, 168, 28, True',
            'ConfigureSafeEditorButton frm.cmdCancel, "Cancel", 634, 448, 90, 28, False',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, module)

        self.assertNotIn("frm.Height = 690", module)

    def test_editor_protects_operational_and_linked_custom_properties(self):
        module = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            "IsCleanupSuiteManagedProperty(CStr(prop.Name))",
            "CBool(prop.LinkToContent)",
            '" linked and "',
            '" CleanupSuite-managed properties remain audit-only."',
            "SafeCustomOriginalNameIsListed",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, module)

    def test_dotm_launch_uses_compile_time_form_reference_and_keeps_real_error(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn('Case "frmMetaDataSuite"', launcher)
        self.assertIn("Set CreateCleanupToolForm = New frmMetaDataSuite", launcher)
        self.assertIn("originalErrNumber = Err.Number", launcher)
        self.assertIn("originalErrDescription = Err.Description", launcher)
        self.assertIn("originalErrSource = Err.Source", launcher)
        self.assertIn("Set mDashboardForm = New frmMetaDataSuite", module)
        self.assertNotIn('VBA.UserForms.Add("frmMetaDataSuite")', module)

    def test_detail_panels_use_capability_wording_not_fake_actions(self):
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn('DASHBOARD_ACTION_EDIT As String = "Available in Safe Edit Current"', module)
        self.assertIn('DASHBOARD_ACTION_SAVE_OPEN As String = "Read-only finding"', module)
        self.assertIn('DASHBOARD_ACTION_DECRYPT As String = "Further investigation available"', module)
        self.assertIn('"[Will clear] Built-in property: "', module)
        self.assertNotIn('DASHBOARD_ACTION_SAVE_OPEN As String = "Save to Open"', module)

    def test_dashboard_smoke_report_passes_required_finding_collections(self):
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn(
            "BuildPackageDetailsText(mDashboardQuickReport.Relationships, _",
            module,
        )
        self.assertIn(
            "BuildCryptoDetailsText(mDashboardQuickReport.CryptoFindings, _",
            module,
        )
        self.assertIn(
            "BuildCryptoDetailsText(mDashboardCryptoFindings, mDashboardGuessFindings)",
            module,
        )
        self.assertNotIn("BuildPackageDetailsText(mDashboardQuickReport)", module)
        self.assertNotIn("BuildCryptoDetailsText()", module)

    def test_detail_panel_property_capabilities_match_safe_editor_scope(self):
        module = read("src/modules/modMetaDataSuite.bas")

        self.assertIn("Private Function BuildDocumentPropertyRows", module)
        self.assertIn("Private Function DashboardDocumentPropertyCapability", module)
        self.assertIn('Case "title", "subject", "creator", "keywords", "description", "category", _', module)
        self.assertIn('"lastmodifiedby", "contentstatus", "contenttype", "language"', module)
        self.assertIn('Case "company", "manager"', module)
        self.assertIn("IsCleanupSuiteManagedProperty(propertyName)", module)
        self.assertIn('"docProps/core.xml", "core"', module)
        self.assertIn('"docProps/app.xml", "application"', module)
        self.assertIn('"docProps/custom.xml", "custom"', module)

    def test_installer_and_sync_create_safe_editor_controls(self):
        installer = read("src/installer/installer.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for source in [installer, sync_script]:
            for snippet in [
                '"frmSafeMetadataEditor"',
                '"lstCustomProperties"',
                '"cboCustomType"',
                '"cmdCustomAddUpdate"',
                '"cmdSave"',
            ]:
                with self.subTest(snippet=snippet):
                    self.assertIn(snippet, source)

        self.assertIn('Forms.ListBox.1', installer)
        self.assertIn('Forms.ComboBox.1', installer)
        self.assertIn('Forms.ListBox.1', sync_script)
        self.assertIn('Forms.ComboBox.1', sync_script)


if __name__ == "__main__":
    unittest.main()
