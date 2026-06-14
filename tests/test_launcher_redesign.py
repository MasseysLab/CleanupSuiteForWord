from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class LauncherRedesignTests(unittest.TestCase):
    def test_launcher_has_category_labels_and_tool_descriptions(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix.",
            installer,
        )
        self.assertNotIn("not by how the VBA code is built", installer)

        for category in [
            "Text and Characters",
            "Paragraphs, Breaks, and Lists",
            "Layout and Document Structure",
            "Formatting, Links, and Styles",
            "Review, Privacy, and Removals",
        ]:
            self.assertIn(category, installer)

        for label in [
            "lblDescUnicode",
            "lblDescDuplicate",
            "lblDescTableClean",
            "lblDescMetadata",
            "lblDescObjectRemover",
        ]:
            self.assertIn(label, installer)

    def test_launcher_layout_is_two_column_row_based_instead_of_flat_button_grid(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("LayoutLauncherCategory", installer)
        self.assertIn("LayoutLauncherToolRow", installer)
        self.assertIn("COL1_X As Single", installer)
        self.assertIn("COL2_X As Single", installer)
        self.assertNotIn("COL3_X As Single", installer)
        self.assertNotIn('Array("cmdUnicode", "cmdHelpUnicode")', installer)

    def test_launcher_is_two_column_with_review_under_right_column(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", COL1_X, yCol1, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            '"lblCatText", "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode"',
            installer,
        )
        self.assertIn("y = yCol1 + 8", installer)
        self.assertLess(
            installer.index("y = yCol1 + 8"),
            installer.index('PositionControl designer, "chkShowCompletionReviewAfterApply"'),
        )

    def test_launcher_title_subtitle_and_close_return_setting_are_wired(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("ApplyMSFormTitleStrategy Me, True", launcher)
        self.assertIn('lblLauncherTitle.Caption = "Cleanup Suite"', launcher)
        self.assertIn("lblLauncherTitle.TextAlign = 2", launcher)
        self.assertIn("lblLauncherSubtitle.TextAlign = 2", launcher)
        self.assertIn(
            'lblLauncherSubtitle.Caption = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."',
            launcher,
        )
        self.assertIn(
            'chkReturnToMainAfterClose.Caption = "Return to main menu after closing individual tool menus"',
            launcher,
        )
        self.assertIn("chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()", launcher)
        self.assertIn("Private Sub chkReturnToMainAfterClose_Click()", launcher)
        self.assertIn("SetReturnToMainAfterCloseSetting chkReturnToMainAfterClose.Value", launcher)
        self.assertIn(
            'chkShowCompletionReviewAfterApply.Caption = "Show completion review after apply"',
            launcher,
        )
        self.assertIn("chkShowCompletionReviewAfterApply.Value = GetShowCompletionReviewAfterApplySetting()", launcher)
        self.assertIn("Private Sub chkShowCompletionReviewAfterApply_Click()", launcher)
        self.assertIn("SetShowCompletionReviewAfterApplySetting chkShowCompletionReviewAfterApply.Value", launcher)
        self.assertIn("Private Sub RefreshApplyReturnCaption()", launcher)
        self.assertIn('chkReturnToMainAfterApply.Caption = "Return to main menu after completion review"', launcher)

        self.assertIn("Public Function GetReturnToMainAfterCloseSetting() As Boolean", helpers)
        self.assertIn("Public Sub SetReturnToMainAfterCloseSetting(enabled As Boolean)", helpers)
        self.assertIn("Public Function GetShowCompletionReviewAfterApplySetting() As Boolean", helpers)
        self.assertIn("Public Sub SetShowCompletionReviewAfterApplySetting(enabled As Boolean)", helpers)
        self.assertIn("Public Sub BeginCleanupToolSession()", helpers)
        self.assertIn("Public Sub MarkCleanupToolApplied()", helpers)
        self.assertIn("Public Sub MarkCleanupToolClosedByUser()", helpers)
        self.assertIn("Public Function ShouldReturnToLauncherAfterToolSession() As Boolean", helpers)
        self.assertIn("Public Sub ReturnToMainAfterToolClose(ByVal closingForm As Object, ByRef Cancel As Integer, ByVal CloseMode As Integer)", helpers)
        self.assertIn("Public Sub ApplyMSFormTitleStrategy", helpers)
        self.assertIn("Private Sub HideLegacyTitleControls", helpers)
        self.assertNotIn("If Not closingForm Is Nothing Then closingForm.Hide", helpers)
        close_helper = helpers[
            helpers.index("Public Sub ReturnToMainAfterToolClose"):
            helpers.index("Public Sub ResetCleanupSuiteDefaults()")
        ]
        self.assertIn("If CloseMode = vbFormControlMenu Then", close_helper)
        self.assertIn("MarkCleanupToolClosedByUser", close_helper)
        self.assertNotIn("ShowCleanupSuiteLauncher", close_helper)
        self.assertIn("SetReturnToMainAfterCloseSetting True", helpers)
        self.assertIn("SetShowCompletionReviewAfterApplySetting True", helpers)

        self.assertIn('"chkShowCompletionReviewAfterApply", "chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave"', installer)
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkReturnToMainAfterClose": ControlCaptionText = "Return to main menu after closing individual tool menus"',
            installer,
        )
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkShowCompletionReviewAfterApply": ControlCaptionText = "Show completion review after apply"',
            installer,
        )
        self.assertIn('ctl.TextAlign = 2', installer)
        self.assertIn('PositionControl designer, "chkReturnToMainAfterClose"', installer)
        self.assertIn('PositionControl designer, "chkReturnToMainAfterApply", MARGIN + 2, y, COL_W, 18', installer)
        self.assertIn('PositionControl designer, "chkShowCompletionReviewAfterApply", MARGIN + 2, y, COL_W, 18', installer)
        self.assertLess(
            installer.index('PositionControl designer, "chkShowCompletionReviewAfterApply"'),
            installer.index('PositionControl designer, "chkReturnToMainAfterApply"'),
        )
        self.assertIn('PositionControl designer, "chkReturnToMainAfterClose", MARGIN + 2, y, COL_W, 18', installer)
        self.assertIn('PositionControl designer, "chkAutoSave", MARGIN + 2, y, COL_W, 18', installer)
        self.assertNotIn('PositionControl designer, "chkReturnToMainAfterApply", MARGIN + 2, y, FORM_W - (2 * MARGIN), 18', installer)

    def test_individual_tool_forms_return_to_main_menu_when_closed_by_user(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if path.name not in {"frmCleanupSuiteLauncher.bas", "frmPreviewActions.bas"}
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", source)
                self.assertIn("ReturnToMainAfterToolClose Me, Cancel, CloseMode", source)

    def test_launcher_reopens_same_tool_without_showing_already_displayed_form(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn("Private Sub OpenCleanupTool(ByVal formName As String)", launcher)
        self.assertIn("BeginCleanupToolSession", launcher)
        self.assertIn("Set toolForm = VBA.UserForms.Add(formName)", launcher)
        self.assertIn("toolForm.Show", launcher)
        self.assertIn("If ShouldReturnToLauncherAfterToolSession() Then Me.Show", launcher)
        self.assertLess(launcher.index("Set toolForm = VBA.UserForms.Add(formName)"), launcher.index("toolForm.Show"))
        self.assertLess(launcher.index("toolForm.Show"), launcher.index("If ShouldReturnToLauncherAfterToolSession() Then Me.Show"))
        self.assertNotIn("toolForm.Hide", launcher)
        self.assertIn('Private Sub cmdCapitalization_Click(): OpenCleanupTool "frmCapitalizationCleanup": End Sub', launcher)
        self.assertNotIn("Private Sub cmdCapitalization_Click(): Me.Hide: frmCapitalizationCleanup.Show: End Sub", launcher)

    def test_launcher_milestone_version_is_070(self):
        launcher = read("src/modules/modCleanupLauncher.bas")
        versioning = read("VERSIONING.md")

        self.assertIn('Public Const SUITE_VERSION As String = "0.7.4"', launcher)
        self.assertIn("## Current Version\n\n`0.7.4`", versioning)
        self.assertIn("0.7.4", versioning)
        self.assertIn("0.7.5", versioning)


if __name__ == "__main__":
    unittest.main()
