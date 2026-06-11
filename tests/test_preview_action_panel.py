from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class PreviewActionPanelTests(unittest.TestCase):
    def test_preview_action_form_is_assembled_and_installed(self):
        manifest = read("src/manifest.txt")
        installer = read("src/installer/installer.bas")

        self.assertIn("forms/frmPreviewActions.bas  frmPreviewActions_Code", manifest)
        self.assertIn('CreateOrReplaceForm vbProj, "frmPreviewActions", frmPreviewActions_Code', installer)
        self.assertIn('AttemptCreateControls vbProj, "frmPreviewActions"', installer)
        self.assertIn('"frmPreviewActions"', installer)
        self.assertIn('Case "frmPreviewActions": FriendlyFormCaption = "Preview Actions"', installer)

    def test_shared_preview_panel_helper_exists(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Public gPreviewActionPanel As Object", helpers)
        self.assertIn("Public Sub ShowPreviewActions", helpers)
        self.assertIn("gPreviewActionPanel.Show", helpers)
        self.assertNotIn("vbModeless", helpers)
        self.assertIn("Public Sub Configure", panel)
        self.assertIn('CallByName mSourceForm, "RunAfterPreview", VbMethod', panel)
        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", panel)
        self.assertIn("Me.Hide", panel)
        self.assertIn("mSourceForm.Show", panel)
        self.assertIn("ShowCleanupSuiteLauncher", panel)
        self.assertIn('cmdClear.Caption = "Reset Preview"', panel)
        self.assertNotIn("cmdBack", panel)
        self.assertNotIn("cmdClose", panel)

    def test_preview_panel_has_only_decisive_actions(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            'ControlsForForm = Array("lblTitle", "lblSummary", "lblHint", "cmdApply", "cmdClear")',
            installer,
        )
        self.assertNotIn('"cmdBack"', installer)
        self.assertNotIn('"cmdClose"', installer)

    def test_launcher_hides_while_subtool_is_open(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn("Private Sub OpenCleanupTool(toolForm As Object)", launcher)
        self.assertIn("Me.Hide", launcher)
        self.assertIn("toolForm.Show", launcher)
        self.assertIn("OpenCleanupTool frmUnicodeCleanup", launcher)
        self.assertNotIn("frmUnicodeCleanup.Show", launcher)

    def test_preview_capable_forms_can_apply_from_panel(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if "chkPreviewOnly" in path.read_text(encoding="utf-8")
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("Public Sub RunAfterPreview()", source)
                self.assertIn("chkPreviewOnly.Value = False", source)
                self.assertIn("cmdRun_Click", source)
                self.assertIn("ShowPreviewActions Me", source)


if __name__ == "__main__":
    unittest.main()
