from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class CapitalizationGuidedFormTests(unittest.TestCase):
    def test_generic_tool_forms_use_shared_guided_layout(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if path.name not in {
                "frmCleanupSuiteLauncher.bas",
                "frmPreviewActions.bas",
                "frmCapitalizationCleanup.bas",
            }
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("LayoutCleanupToolForm Me", source)
                self.assertNotIn("LayoutCapitalizationCleanupForm Me", source)

        self.assertIn("Private Function EnsureGuidedLabel(ByVal toolForm As Object, ByVal controlName As String) As Object", helpers)
        self.assertIn('Set titleLabel = EnsureGuidedLabel(toolForm, "lblGuidedToolName")', helpers)
        self.assertIn('Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")', helpers)
        self.assertIn("ApplyMSFormTitleStrategy toolForm, True", helpers)
        self.assertIn('If Left$(ctl.Name, 3) = "fra" Then', helpers)
        self.assertIn("LayoutGuidedActionButtons toolForm, M, y, contentW", helpers)
        self.assertIn("Private Sub LayoutGuidedScopeRow(ByVal toolForm As Object", helpers)
        self.assertIn("Private Function GuidedToolIntro(ByVal formName As String) As String", helpers)

        self.assertIn("Private Sub LayoutGenericToolControls(comp As VBIDE.VBComponent, designer As Object)", installer)
        self.assertIn("LayoutGenericToolControls comp, designer", installer)
        self.assertIn('If Left$(CStr(cn), 3) = "fra" Then', installer)
        self.assertIn('PositionControl designer, "cmdPreview", MARGIN, y, contentW, BTN_H', installer)
        self.assertIn('PositionControl designer, "cmdRun", MARGIN, y, thirdW, BTN_H', installer)
        self.assertIn('PositionControl designer, "cmdHelp", MARGIN + (2 * (thirdW + GAP)), y, thirdW, BTN_H', installer)

    def test_capitalization_form_uses_guided_layout_without_behavior_change(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('lblIntro.Caption = "Choose how you want capitalization cleaned up."', form)
        self.assertIn("ApplyMSFormTitleStrategy Me, True", form)
        self.assertIn('optAll.Caption = "Fix sentence starts"', form)
        self.assertIn('optUpper.Caption = "UPPERCASE"', form)
        self.assertIn('optLower.Caption = "lowercase"', form)
        self.assertIn('optCustom.Caption = "Custom"', form)
        self.assertIn("Private Sub UpdateModeSummary()", form)
        self.assertIn("Recommended: capitalizes likely sentence starts", form)
        self.assertIn("Custom runs selected choices in order", form)
        self.assertNotIn("the last\" & vbCrLf", form)
        self.assertNotIn("Help  --  Capitalization Fixer", form)
        self.assertIn('ShowCleanupToolHelp "Capitalization"', form)
        self.assertIn("Private Sub RefreshCapitalizationChoices()", form)
        self.assertIn("chkSentence.Visible = optCustom.Value", form)
        self.assertIn("chkSmartSentences.Visible = optCustom.Value", form)
        self.assertIn("fraCustom.Visible = False", form)
        self.assertIn("LayoutCapitalizationCleanupForm Me", form)
        self.assertNotIn("LayoutCleanupToolForm Me", form)

        self.assertIn("If optAll.Value Then doSmart = True", form)
        self.assertIn("If optSentence.Value Then doSentence = True", form)
        self.assertIn("If optTitle.Value Then doTitle = True", form)
        self.assertIn("If optUpper.Value Then doUpper = True", form)
        self.assertIn("If optLower.Value Then doLower = True", form)
        self.assertIn(
            "If optCustom.Value Then doSentence = chkSentence.Value: doTitle = chkTitle.Value: doUpper = chkUpper.Value: doLower = chkLower.Value: doSmart = chkSmartSentences.Value",
            form,
        )

    def test_capitalization_guided_controls_are_generated(self):
        installer = read("src/installer/installer.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('"lblIntro", "optAll", "optSentence"', installer)
        self.assertIn('"optCustom", "lblModeSummary", "fraCustom"', installer)
        self.assertIn(
            'Case "frmCapitalizationCleanup.lblIntro": ControlCaptionText = "Choose how you want capitalization cleaned up."',
            installer,
        )
        self.assertIn('Case "frmCapitalizationCleanup.optAll": ControlCaptionText = "Fix sentence starts"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optUpper": ControlCaptionText = "UPPERCASE"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optLower": ControlCaptionText = "lowercase"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optCustom": ControlCaptionText = "Custom"', installer)
        self.assertIn("Private Sub LayoutCapitalizationControls(comp As VBIDE.VBComponent, designer As Object)", installer)
        self.assertIn("Private Function FormBodyOwnsTitle(formName As String) As Boolean", installer)
        self.assertIn('Case "frmCleanupSuiteLauncher", "frmPreviewActions", "frmCapitalizationCleanup"', installer)
        self.assertIn('designer.Caption = ""', installer)
        self.assertIn('If comp.Name = "frmCapitalizationCleanup" Then', installer)
        self.assertIn('LayoutCapitalizationControls comp, designer', installer)
        self.assertIn('PositionControl designer, "lblIntro"', installer)
        self.assertIn('PositionControl designer, "lblModeSummary"', installer)
        self.assertIn('PositionControl designer, "fraCustom", 0, 0, 0, 0', installer)
        self.assertIn('designer.Controls("fraCustom").Visible = False', installer)
        self.assertIn('PositionControl designer, "chkSentence", MARGIN + 8, y, contentW - 16, 16', installer)

        self.assertIn("Public Sub LayoutCapitalizationCleanupForm(ByVal toolForm As Object)", helpers)
        self.assertIn('toolForm.Controls("chkPreviewOnly").Visible = False', helpers)
        self.assertIn('toolForm.Controls("fraScopeSelection").Visible = False', helpers)
        self.assertIn('PositionGuidedChoice toolForm, "optAll"', helpers)
        self.assertIn('With toolForm.Controls("lblModeSummary")', helpers)
        self.assertIn('With toolForm.Controls("cmdPreview")', helpers)
        self.assertIn('With toolForm.Controls("cmdRun")', helpers)
        self.assertIn("ApplyMSFormTitleStrategy toolForm, True", helpers)
        self.assertIn("Public Sub ApplyMSFormTitleStrategy", helpers)
        self.assertIn("Private Sub HideLegacyTitleControls(ByVal toolForm As Object)", helpers)
        self.assertIn("Private Function IsLegacyTitleControlName(controlName As String) As Boolean", helpers)
        self.assertIn('IsLegacyTitleControlName = (nm = "lbltitle" Or Left$(nm, 8) = "lbltitle")', helpers)
        self.assertIn(".ZOrder 0", helpers)
        self.assertIn("toolForm.Height = y + 34", helpers)
        self.assertIn('toolForm.Controls("fraCustom").Visible = False', helpers)
        self.assertIn('If toolForm.Controls("optCustom").Value Then', helpers)
        self.assertIn('PositionGuidedCheck toolForm, "chkSentence", M + 8, y, contentW - 16', helpers)
        self.assertIn('Private Sub HideGuidedCustom(ByVal toolForm As Object)', helpers)

        self.assertIn('comp.Properties("Height") = y + 34', installer)
        self.assertIn('designer.Height = y + 34', installer)


if __name__ == "__main__":
    unittest.main()
