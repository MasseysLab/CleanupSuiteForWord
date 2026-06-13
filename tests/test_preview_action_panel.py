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
        self.assertIn("Private gPreviewActionPanelHasPosition As Boolean", helpers)
        self.assertIn("Private gPreviewActionPanelLeft As Single", helpers)
        self.assertIn("Private gPreviewActionPanelTop As Single", helpers)
        self.assertIn("Public Sub RememberPreviewActionPanelPosition", helpers)
        self.assertIn("Private Sub ApplyPreviewActionPanelPosition", helpers)
        self.assertIn("gPreviewActionPanelHasPosition = True", helpers)
        self.assertIn("panel.StartUpPosition = 0", helpers)
        self.assertIn("panel.Left = gPreviewActionPanelLeft", helpers)
        self.assertIn("panel.Top = gPreviewActionPanelTop", helpers)
        self.assertIn("Public Sub ShowPreviewActions", helpers)
        self.assertIn("gPreviewActionPanel.Show", helpers)
        self.assertLess(
            helpers.index("ApplyPreviewActionPanelPosition gPreviewActionPanel"),
            helpers.index("gPreviewActionPanel.Show"),
        )
        self.assertNotIn("vbModeless", helpers)
        self.assertIn("Public Sub Configure", panel)
        self.assertIn('CallByName mSourceForm, "RunAfterPreview", VbMethod', panel)
        self.assertIn('CallByName src, "PreviewFromPanel", VbMethod', panel)
        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", panel)
        self.assertIn("Me.Show vbModeless", panel)
        self.assertIn("Me.Hide", panel)
        self.assertIn("mSourceForm.Show", panel)
        self.assertIn("ShowCleanupSuiteLauncher", panel)
        self.assertIn("Private Sub LayoutPanel()", panel)
        self.assertIn("LayoutPanel", panel)
        self.assertIn("Private Sub ShowModelessAtCurrentPosition()", panel)
        self.assertIn("Dim panelLeft As Single", panel)
        self.assertIn("Dim panelTop As Single", panel)
        self.assertIn("panelLeft = Me.Left", panel)
        self.assertIn("panelTop = Me.Top", panel)
        self.assertIn("Me.Left = panelLeft", panel)
        self.assertIn("Me.Top = panelTop", panel)
        self.assertIn("RememberPreviewActionPanelPosition Me.Left, Me.Top", panel)
        self.assertLess(
            panel.index("RememberPreviewActionPanelPosition Me.Left, Me.Top"),
            panel.index('CallByName src, "PreviewFromPanel", VbMethod'),
        )
        self.assertIn("Private Sub StyleActionBar()", panel)
        self.assertIn("Private Sub StyleLabel", panel)
        self.assertIn("Private Sub StyleButton", panel)
        self.assertIn("StyleLabel lblHint, 8, True", panel)
        self.assertIn("StyleLabel lblSummary, 7.5, False", panel)
        self.assertIn("ctl.Font.Size = 7.5", panel)
        self.assertIn("Const FORM_W As Single = 300", panel)
        self.assertIn("Const FORM_H As Single = 90", panel)
        self.assertIn("Const CONTENT_W As Single = 262", panel)
        self.assertIn("Const BTN_W As Single = 84", panel)
        self.assertIn("Const BH As Single = 18", panel)
        self.assertIn("Const GAP As Single = 5", panel)
        self.assertIn('lblTitle.Caption = ""', panel)
        self.assertIn("lblTitle.Visible = False", panel)
        self.assertIn("lblTitle.Move 0, 0, 0, 0", panel)
        self.assertIn("lblHint.Caption = toolName", panel)
        self.assertIn("lblSummary.Caption = mSummaryText", panel)
        self.assertIn("lblHint.Move M, 2, CONTENT_W, 12", panel)
        self.assertIn("cmdPreview.Move M, 17, BTN_W, BH", panel)
        self.assertIn("cmdClear.Move M + BTN_W + GAP, 17, BTN_W, BH", panel)
        self.assertIn("cmdApply.Move M + (2 * (BTN_W + GAP)), 17, BTN_W, BH", panel)
        self.assertIn("lblSummary.Move M, 40, CONTENT_W, 16", panel)
        self.assertIn("If Not DocumentHasVisibleContent() Then", panel)
        self.assertIn("This document appears to be blank. There is nothing to apply.", panel)
        self.assertIn('cmdPreview.Caption = "Preview is ON"', panel)
        self.assertIn("cmdPreview.Enabled = True", panel)
        self.assertIn('cmdPreview.Caption = "Preview is OFF"', panel)
        self.assertIn('lblSummary.Caption = "You may now edit your document if you wish."', panel)
        self.assertLess(
            panel.index('cmdPreview.Caption = "Preview is OFF"'),
            panel.index('lblSummary.Caption = "You may now edit your document if you wish."'),
        )
        self.assertIn("ShowModelessAtCurrentPosition", panel)
        self.assertIn('cmdClear.Caption = "Reconfigure"', panel)
        self.assertNotIn('cmdClear.Caption = "Settings"', panel)
        self.assertNotIn('cmdClear.Caption = "Reset Preview"', panel)
        self.assertNotIn("UserForm_Initialize", panel)
        self.assertNotIn("cmdBack", panel)
        self.assertNotIn("cmdClose", panel)

    def test_preview_panel_has_only_decisive_actions(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            'ControlsForForm = Array("lblTitle", "lblSummary", "lblHint", "cmdApply", "cmdPreview", "cmdClear")',
            installer,
        )
        self.assertNotIn('"cmdBack"', installer)
        self.assertNotIn('"cmdClose"', installer)

    def test_launcher_hides_while_subtool_is_open(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertNotIn("OpenCleanupTool(toolForm As Object)", launcher)
        self.assertNotIn("OpenCleanupTool frmUnicodeCleanup", launcher)
        self.assertIn("Private Sub cmdUnicode_Click(): Me.Hide: frmUnicodeCleanup.Show: End Sub", launcher)
        self.assertIn("Private Sub cmdCapitalization_Click(): Me.Hide: frmCapitalizationCleanup.Show: End Sub", launcher)

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

    def test_apply_and_preview_button_language(self):
        installer = read("src/installer/installer.bas")

        self.assertIn('If CStr(cn) = "cmdPreview" Then', installer)
        self.assertIn('PositionControl designer, "cmdPreview", MARGIN, y, contentW, BTN_H', installer)
        self.assertIn("Const FORM_W As Single = 300", installer)
        self.assertIn("Const FORM_H As Single = 90", installer)
        self.assertIn("Const CONTENT_W As Single = 262", installer)
        self.assertIn("Const BTN_W As Single = 84", installer)
        self.assertIn("Const BTN_H As Single = 18", installer)
        self.assertIn("Const GAP As Single = 5", installer)
        self.assertIn('Case "frmPreviewActions.lblTitle": ControlCaptionText = ""', installer)
        self.assertIn('Case "frmPreviewActions.lblHint": ControlCaptionText = "Tool Name"', installer)
        self.assertIn('comp.Properties("Caption") = ""', installer)
        self.assertIn('designer.Caption = ""', installer)
        self.assertIn('Case "cmdClear": cap = "Reconfigure"', installer)
        self.assertIn('Case "cmdRun": cap = "Apply"', installer)
        self.assertIn('Case "cmdPreview": cap = "Preview"', installer)
        self.assertIn('Case "cmdReset": cap = "Reset"', installer)
        self.assertIn('Case "cmdResetAll": cap = "Reset All"', installer)
        self.assertIn('Case "cmdPreview": ButtonTipText = "Preview with the current settings"', installer)
        self.assertIn('Case "cmdReset": ButtonTipText = "Reset this tool to its defaults"', installer)
        self.assertIn('Case "cmdResetAll": ButtonTipText = "Reset all tools and global settings to defaults"', installer)
        self.assertIn("ctl.Enabled = True", installer)
        self.assertIn('ControlTypeByName = "Forms.CommandButton.1"', installer)
        self.assertNotIn('"chkPreviewOnly", "cmdRun"', installer)
        self.assertNotIn('"cmdDeselectAll", "cmdRun"', installer)

    def test_blank_document_apply_is_short_circuited_before_tool_runs(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Public Function DocumentHasVisibleContent", helpers)
        self.assertIn('textValue = Replace(textValue, vbCr, "")', helpers)
        self.assertIn("ActiveDocument.InlineShapes.Count", helpers)
        self.assertIn("ActiveDocument.Shapes.Count", helpers)
        self.assertIn("ActiveDocument.Tables.Count", helpers)
        self.assertIn("If Not DocumentHasVisibleContent() Then", panel)
        self.assertLess(
            panel.index("If Not DocumentHasVisibleContent() Then"),
            panel.index('CallByName mSourceForm, "RunAfterPreview", VbMethod'),
        )

    def test_preview_capable_forms_have_preview_button_hook(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if "chkPreviewOnly" in path.read_text(encoding="utf-8")
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("Private Sub cmdPreview_Click()", source)
                self.assertIn("Public Sub PreviewFromPanel()", source)
                self.assertIn("Private Sub cmdReset_Click()", source)
                self.assertIn("UserForm_Initialize", source)
                self.assertIn("LayoutCleanupToolForm Me", source)
                self.assertIn("chkPreviewOnly.Value = True", source)
                self.assertIn("cmdRun_Click", source)

    def test_return_to_main_after_apply_setting_is_wired(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()", launcher)
        self.assertIn("SetReturnToMainAfterApplySetting chkReturnToMainAfterApply.Value", launcher)
        self.assertIn("Public Function GetReturnToMainAfterApplySetting() As Boolean", helpers)
        self.assertIn("Public Sub SetReturnToMainAfterApplySetting(enabled As Boolean)", helpers)
        self.assertIn("Public Sub ResetCleanupSuiteDefaults()", helpers)
        self.assertIn("SetAutoSaveSetting True", helpers)
        self.assertIn("SetReturnToMainAfterApplySetting True", helpers)
        self.assertIn("If GetReturnToMainAfterApplySetting() Then ShowCleanupSuiteLauncher", panel)
        self.assertIn('"chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave", "cmdResetAll"', installer)
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkReturnToMainAfterApply": ControlCaptionText = "Return to main menu after apply"',
            installer,
        )
        self.assertIn("Private Sub cmdResetAll_Click()", launcher)
        self.assertIn("ResetCleanupSuiteDefaults", launcher)

    def test_installer_report_write_is_unique_and_nonfatal(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("WriteInstallReport reportText", installer)
        self.assertIn("Private Sub WriteInstallReport(reportText As String)", installer)
        self.assertIn('CleanupSuite_Install_Report_" & Format$(Now, "yyyymmdd_hhnnss")', installer)
        self.assertIn("ReportErr:", installer)
        self.assertIn("Installer finished, but the report could not be written or opened.", installer)
        self.assertIn("The Cleanup Suite components were already installed before the report step.", installer)
        self.assertNotIn('reportPath = Environ$("TEMP") & "\\CleanupSuite_Install_Report.txt"', installer)


if __name__ == "__main__":
    unittest.main()
