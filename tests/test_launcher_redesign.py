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
            "Clean Pasted Text",
            "Fix Paragraphs and Lists",
            "Clean Document Objects",
            "Clean Formatting",
            "Review and Inspect",
            "Prepare for Sharing",
            "Risky Alpha Tools",
        ]:
            self.assertIn(category, installer)

        for label in [
            "lblDescUnicode",
            "lblDescDuplicate",
            "lblDescTableClean",
            "lblDescMetadata",
            "lblDescFinalReview",
            "lblDescObjectRemover",
        ]:
            self.assertIn(label, installer)

    def test_launcher_has_risk_labels_and_settled_first_pass_names(self):
        installer = read("src/installer/installer.bas")
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        for control in [
            "lblRiskUnicode",
            "lblRiskPunctuation",
            "lblRiskSpacing",
            "lblRiskCapitalization",
            "lblRiskList",
            "lblRiskParagraph",
            "lblRiskSoftReturn",
            "lblRiskDuplicate",
            "lblRiskTableClean",
            "lblRiskBreakNorm",
            "lblRiskHeaderFooter",
            "lblRiskDocTrim",
            "lblRiskFontNorm",
            "lblRiskFormatStrip",
            "lblRiskStyleClean",
            "lblRiskHyperlink",
            "lblRiskMetadata",
            "lblRiskFinalReview",
            "lblRiskFootnote",
            "lblRiskObjectRemover",
        ]:
            with self.subTest(control=control):
                self.assertIn(control, installer)
                self.assertIn(control, sync_script)

        for label in [
            "Safe cleanup",
            "Text caution",
            "Structure change",
            "Inspect First",
            "Formatting change",
            "Alpha tool",
            "Removes content",
            "Privacy cleanup",
            "Finalizes review",
        ]:
            with self.subTest(label=label):
                self.assertIn(label, installer)

        self.assertIn('Case "cmdDocTrim": cap = "Trim End of Doc"', installer)
        self.assertIn('Case "frmCleanupSuiteLauncher.lblDescSoftReturn": ControlCaptionText = "Convert Shift+Enter line breaks to paragraph marks, or the reverse of that."', installer)
        self.assertNotIn("or reverse that conversion.", installer)
        self.assertIn('Case "cmdFormatStrip": cap = "Formatting Cleaner"', installer)
        self.assertIn('Case "frmCleanupSuiteLauncher.lblDescDuplicate": ControlCaptionText = "Find repeated or near-duplicate paragraphs before removing them."', installer)
        self.assertIn('Case "frmCleanupSuiteLauncher.lblDescHyperlink": ControlCaptionText = "Keep links but Hide styling" & vbCrLf & "Keep text but remove links"', installer)
        self.assertIn('lblDescHyperlink.Caption = "Keep links but Hide styling" & vbCrLf & "Keep text but remove links"', launcher)
        self.assertNotIn("Hide hyperlink styling while keeping links, or remove link targets while keeping visible text.", installer)
        self.assertNotIn("Hide hyperlink styling while keeping links, or remove link targets while keeping visible text.", launcher)
        self.assertIn('cmdDocTrim.Caption = "Trim End of Doc"', launcher)
        self.assertIn('cmdFormatStrip.Caption = "Formatting Cleaner"', launcher)
        self.assertNotIn('buttonCtl.Name = "cmdDocTrim"', launcher)
        self.assertNotIn('buttonName = "cmdDocTrim"', installer)

    def test_launcher_uses_seven_intent_groups_and_chip_colors(self):
        installer = read("src/installer/installer.bas")
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        for group in [
            "Clean Pasted Text",
            "Fix Paragraphs and Lists",
            "Clean Formatting",
            "Clean Document Objects",
            "Review and Inspect",
            "Prepare for Sharing",
            "Risky Alpha Tools",
        ]:
            with self.subTest(group=group):
                self.assertIn(group, installer)

        self.assertIn("StyleLauncherRiskLabel", installer)
        self.assertIn("LauncherRiskBackColor", installer)
        self.assertIn("LauncherRiskForeColor", installer)
        self.assertIn("riskLabel.TakeFocusOnClick = False", launcher)
        self.assertIn('If Left$(ctrlName, 7) = "lblRisk" Then', installer)
        self.assertIn('ControlTypeByName = "Forms.CommandButton.1"', installer)
        self.assertIn(".TakeFocusOnClick = False", installer)
        self.assertIn('LayoutLauncherToolRow(designer, "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "lblRiskUnicode", COL1_X, yCol1, COL_W)', installer)
        self.assertIn('LayoutLauncherToolRow(designer, "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", "lblRiskMetadata", COL2_X, yCol2, COL_W)', installer)
        self.assertIn('LayoutLauncherToolRow(designer, "cmdHelpFinalReview", "cmdFinalReview", "lblDescFinalReview", "lblRiskFinalReview", COL2_X, yCol2, COL_W)', installer)
        self.assertNotIn("Fix Paragraphs && Lists", installer)
        self.assertNotIn("Review && Inspect", installer)

    def test_launcher_layout_is_two_column_row_based_instead_of_flat_button_grid(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("LayoutLauncherCategory", installer)
        self.assertIn("LayoutLauncherToolRow", installer)
        self.assertIn("COL1_X As Single", installer)
        self.assertIn("COL2_X As Single", installer)
        self.assertNotIn("COL3_X As Single", installer)
        self.assertNotIn('Array("cmdUnicode", "cmdHelpUnicode")', installer)
        self.assertIn("Const ROW_H As Single = 30", installer)
        self.assertIn("Const BTN_W As Single = 108", installer)
        self.assertIn("Const RISK_W As Single = 56", installer)
        self.assertIn("Const RISK_H As Single = 28", installer)
        self.assertIn("T + 1, RISK_W, RISK_H", installer)
        self.assertIn("calloutWidth = (MARGIN + COL_W) - calloutLeft", installer)
        self.assertIn("calloutHeight = 32", installer)
        self.assertIn('PositionControl designer, "lblResetAllCalloutBox", calloutLeft, calloutTop, calloutWidth, calloutHeight', installer)
        self.assertIn('PositionControl designer, "lblResetAllCalloutText", calloutLeft + 5, calloutTop + 4, calloutWidth - 10, calloutHeight - 8', installer)
        self.assertIn(".WordWrap = True", installer)
        self.assertIn('LauncherRiskDisplayText = Replace(labelText, " ", vbCrLf)', installer)
        self.assertIn('.Caption = LauncherRiskDisplayText(labelText)', installer)
        self.assertIn("If highlights remain after an error, select Reset All.", installer)
        self.assertNotIn("then select Reset All, and it should remove them.", installer)

    def test_launcher_is_two_column_with_review_under_right_column(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "lblRiskUnicode", COL1_X, yCol1, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", "lblRiskMetadata", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpFinalReview", "cmdFinalReview", "lblDescFinalReview", "lblRiskFinalReview", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherCategory(designer, "lblCatReview", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpDuplicate", "cmdDuplicate", "lblDescDuplicate", "lblRiskDuplicate", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherCategory(designer, "lblCatAlpha", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            'LayoutLauncherToolRow(designer, "cmdHelpStyle", "cmdStyleClean", "lblDescStyleClean", "lblRiskStyleClean", COL2_X, yCol2, COL_W)',
            installer,
        )
        self.assertIn(
            '"lblCatText", "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "lblRiskUnicode"',
            installer,
        )
        self.assertIn("y = yCol1 + GAP", installer)
        self.assertLess(
            installer.index("y = yCol1 + GAP"),
            installer.index('PositionControl designer, "chkReturnToMainAfterApply"'),
        )
        self.assertLess(
            installer.index('LayoutLauncherCategory(designer, "lblCatReview", COL2_X, yCol2, COL_W)'),
            installer.index('LayoutLauncherCategory(designer, "lblCatShare", COL2_X, yCol2, COL_W)'),
        )
        self.assertLess(
            installer.index('LayoutLauncherCategory(designer, "lblCatShare", COL2_X, yCol2, COL_W)'),
            installer.index('LayoutLauncherCategory(designer, "lblCatAlpha", COL2_X, yCol2, COL_W)'),
        )

    def test_launcher_title_subtitle_and_close_return_setting_are_wired(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn("ApplyMSFormTitleStrategy Me, True", launcher)
        self.assertIn('lblLauncherTitle.Caption = "Cleanup Suite"', launcher)
        self.assertIn("lblLauncherTitle.TextAlign = 1", launcher)
        self.assertIn("lblLauncherSubtitle.TextAlign = 1", launcher)
        self.assertIn(
            'lblLauncherSubtitle.Caption = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."',
            launcher,
        )
        self.assertIn("RefreshReviewPrivacyLauncherSection", launcher)
        self.assertIn(
            'chkReturnToMainAfterClose.Caption = "Global default: return to main menu after closing individual tool menus"',
            launcher,
        )
        self.assertIn("chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()", launcher)
        self.assertIn("Private Sub chkReturnToMainAfterClose_Click()", launcher)
        self.assertIn("SetReturnToMainAfterCloseSetting chkReturnToMainAfterClose.Value", launcher)
        self.assertNotIn("chkShowCompletionReviewAfterApply", launcher)
        self.assertNotIn("GetShowCompletionReviewAfterApplySetting", launcher)
        self.assertNotIn("SetShowCompletionReviewAfterApplySetting", launcher)
        self.assertNotIn("Private Sub RefreshApplyReturnCaption()", launcher)
        self.assertIn('chkReturnToMainAfterApply.Caption = "Global default: return to main menu after apply"', launcher)

        self.assertIn("Public Function GetReturnToMainAfterCloseSetting() As Boolean", helpers)
        self.assertIn("Public Sub SetReturnToMainAfterCloseSetting(enabled As Boolean)", helpers)
        self.assertNotIn("GetShowCompletionReviewAfterApplySetting", helpers)
        self.assertNotIn("SetShowCompletionReviewAfterApplySetting", helpers)
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
            helpers.index("Public Sub ReturnToLauncherAfterDetachedToolSession()")
        ]
        self.assertIn("If CloseMode = vbFormControlMenu Then", close_helper)
        self.assertIn("MarkCleanupToolClosedByUser", close_helper)
        self.assertNotIn("ShowCleanupSuiteLauncher", close_helper)
        self.assertIn("Public Sub ReturnToLauncherAfterDetachedToolSession()", helpers)
        self.assertIn("SetReturnToMainAfterCloseSetting True", helpers)
        self.assertNotIn("SetShowCompletionReviewAfterApplySetting True", helpers)

        self.assertIn('"chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave"', installer)
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkReturnToMainAfterClose": ControlCaptionText = "Global default: return to main menu after closing individual tool menus"',
            installer,
        )
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkReturnToMainAfterApply": ControlCaptionText = "Global default: return to main menu after apply"',
            installer,
        )
        self.assertIn('ctl.TextAlign = 1', installer)
        self.assertIn('PositionControl designer, "chkReturnToMainAfterClose"', installer)
        self.assertIn('PositionControl designer, "chkReturnToMainAfterApply", MARGIN + 2, y, COL_W, 18', installer)
        self.assertNotIn('PositionControl designer, "chkShowCompletionReviewAfterApply"', installer)
        self.assertIn('PositionControl designer, "chkReturnToMainAfterClose", MARGIN + 2, y, COL_W, 18', installer)
        self.assertIn('PositionControl designer, "chkAutoSave", MARGIN + 2, y, 225, 18', installer)
        self.assertIn('PositionControl designer, "lblAutoSaveWarning", MARGIN + 18, y + 16, COL_W - 18, 14', installer)
        self.assertIn('PositionLauncherControl lblAutoSaveWarning, MARGIN + 18, y + 16, COL_W - 18, 14', launcher)
        self.assertNotIn('PositionControl designer, "lblAutoSaveWarning", MARGIN + 232, y, COL_W - 232, 18', installer)
        self.assertIn('lblAutoSaveWarning.Caption = "It is dangerous to turn off auto save!"', launcher)
        self.assertIn('launcher.Controls("lblAutoSaveWarning").Top > launcher.Controls("chkAutoSave").Top', smoke)
        self.assertIn('.Visible = (Not chkAutoSave.Value)', launcher)
        self.assertNotIn('PositionControl designer, "chkReturnToMainAfterApply", MARGIN + 2, y, FORM_W - (2 * MARGIN), 18', installer)

    def test_individual_tool_forms_return_to_main_menu_when_closed_by_user(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if path.name not in {"frmCleanupSuiteLauncher.bas", "frmPreviewActions.bas", "frmSafeMetadataEditor.bas"}
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
        self.assertIn("Set toolForm = CreateCleanupToolForm(formName)", launcher)
        self.assertIn("Set CreateCleanupToolForm = VBA.UserForms.Add(formName)", launcher)
        self.assertIn("toolForm.Show", launcher)
        self.assertIn("If ShouldReturnToLauncherAfterToolSession() Then Me.Show", launcher)
        self.assertLess(launcher.index("Set toolForm = CreateCleanupToolForm(formName)"), launcher.index("toolForm.Show"))
        self.assertLess(launcher.index("toolForm.Show"), launcher.index("If ShouldReturnToLauncherAfterToolSession() Then Me.Show"))
        self.assertNotIn("toolForm.Hide", launcher)
        self.assertIn('Private Sub cmdCapitalization_Click(): OpenCleanupTool "frmCapitalizationCleanup": End Sub', launcher)
        self.assertNotIn("Private Sub cmdCapitalization_Click(): Me.Hide: frmCapitalizationCleanup.Show: End Sub", launcher)

    def test_launcher_milestone_version_is_092_alpha(self):
        launcher = read("src/modules/modCleanupLauncher.bas")
        versioning = read("VERSIONING.md")
        launcher_form = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn('Public Const SUITE_VERSION As String = "0.9.3-beta-vba-final"', launcher)
        self.assertIn("## Current Version\n\n`0.9.3-beta-vba-final`", versioning)
        self.assertIn("0.7.4", versioning)
        self.assertIn("0.7.4.5", versioning)
        self.assertIn("0.7.5", versioning)
        self.assertIn("0.8.0", versioning)
        self.assertIn('cmdFootnote.Caption = "Footnote / Endnote"', launcher_form)
        self.assertIn('cmdMetadata.Caption = "MetaDataSuite"', launcher_form)
        self.assertNotIn("chkEnableAdvancedMetadataTools", launcher_form)
        self.assertIn('cmdFinalReview.Caption = "Final Review"', launcher_form)

    def test_review_privacy_refresh_normalizes_final_review_button_and_help_style(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn('cmdHelpFinalReview.Caption = "?"', launcher)
        self.assertIn("cmdHelpFinalReview.Font.Bold = True", launcher)
        self.assertIn("cmdHelpFinalReview.BackColor = cmdHelpMetadata.BackColor", launcher)
        self.assertIn("cmdHelpFinalReview.ForeColor = cmdHelpMetadata.ForeColor", launcher)
        self.assertIn("cmdFinalReview.Font.Bold = True", launcher)
        self.assertIn("cmdFinalReview.BackColor = cmdMetadata.BackColor", launcher)
        self.assertIn("cmdFinalReview.ForeColor = cmdMetadata.ForeColor", launcher)
        self.assertIn("lblDescFinalReview.Font.Bold = False", launcher)
        self.assertIn("lblDescFinalReview.Font.Size = lblDescMetadata.Font.Size", launcher)
        self.assertIn("lblDescFinalReview.ForeColor = lblDescMetadata.ForeColor", launcher)
        self.assertIn("lblDescFinalReview.BackColor = lblDescMetadata.BackColor", launcher)

    def test_launcher_runtime_normalizes_user_manual_button_caption_and_style(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn('cmdUserManual.Caption = "User Manual"', launcher)
        self.assertIn("cmdUserManual.Font.Bold = cmdResetAll.Font.Bold", launcher)
        self.assertIn("cmdUserManual.BackColor = cmdResetAll.BackColor", launcher)
        self.assertIn("cmdUserManual.ForeColor = cmdResetAll.ForeColor", launcher)

    def test_launcher_risk_chips_explain_their_label_when_clicked(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn("Private Sub ShowLauncherRiskExplanation", launcher)
        self.assertIn('MsgBox msg, vbInformation, "Why This Risk?"', launcher)
        self.assertIn('"Risk label: " & riskLabel', launcher)
        self.assertIn('"Why this tool says that:"', launcher)

        for control_name in [
            "lblRiskUnicode",
            "lblRiskPunctuation",
            "lblRiskSpacing",
            "lblRiskCapitalization",
            "lblRiskList",
            "lblRiskParagraph",
            "lblRiskSoftReturn",
            "lblRiskDocTrim",
            "lblRiskDuplicate",
            "lblRiskTableClean",
            "lblRiskHeaderFooter",
            "lblRiskFootnote",
            "lblRiskObjectRemover",
            "lblRiskFontNorm",
            "lblRiskFormatStrip",
            "lblRiskMetadata",
            "lblRiskFinalReview",
            "lblRiskHyperlink",
            "lblRiskStyleClean",
            "lblRiskBreakNorm",
        ]:
            with self.subTest(control=control_name):
                self.assertIn(f"Private Sub {control_name}_Click()", launcher)

        for label in [
            "Safe cleanup",
            "Text caution",
            "Structure change",
            "Inspect First",
            "Formatting change",
            "Privacy cleanup",
            "Finalizes review",
            "Removes content",
            "Alpha tool",
        ]:
            with self.subTest(label=label):
                self.assertIn(f', "{label}", ', launcher)

    def test_launcher_has_reset_all_highlight_recovery_callout(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        installer = read("src/installer/installer.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")
        callout_text = "If highlights remain after an error, select Reset All."

        self.assertIn("RefreshResetAllCallout", launcher)
        self.assertIn("lblResetAllArrow.Caption = ChrW$(&H2190)", launcher)
        self.assertIn(f'lblResetAllCalloutText.Caption = "{callout_text}"', launcher)
        self.assertIn("lblResetAllCalloutBox.BorderStyle = 1", launcher)
        self.assertIn("lblResetAllCalloutBox.ForeColor = RGB(220, 0, 0)", launcher)
        self.assertIn("lblResetAllArrow.ForeColor = RGB(220, 0, 0)", launcher)

        for control_name in [
            "lblResetAllArrow",
            "lblResetAllCalloutBox",
            "lblResetAllCalloutText",
        ]:
            self.assertIn(control_name, installer)
            self.assertIn(control_name, sync_script)

        self.assertIn(
            f'Case "frmCleanupSuiteLauncher.lblResetAllCalloutText": ControlCaptionText = "{callout_text}"',
            installer,
        )
        self.assertIn(
            'PositionControl designer, "lblResetAllCalloutBox"',
            installer,
        )
        self.assertIn(
            'PositionControl designer, "lblResetAllCalloutText"',
            installer,
        )


if __name__ == "__main__":
    unittest.main()
