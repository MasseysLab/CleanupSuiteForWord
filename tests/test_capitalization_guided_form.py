from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class ToolFormRulesTests(unittest.TestCase):
    def test_tool_forms_use_shared_guided_layout_without_help_handlers(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if path.name not in {"frmCleanupSuiteLauncher.bas", "frmPreviewActions.bas"}
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("LayoutCleanupToolForm Me", source)
                self.assertNotIn("LayoutCapitalizationCleanupForm Me", source)
                self.assertNotIn("cmdHelp_Click", source)
                self.assertNotIn("ShowCleanupToolHelp", source)

    def test_shared_layout_matches_tool_form_rules(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")', helpers)
        self.assertIn("ApplyMSFormTitleStrategy toolForm, True", helpers)
        self.assertIn('RemoveGuidedGeneratedLabel toolForm, "lblGuidedToolName"', helpers)
        self.assertIn(".Font.Bold = True", helpers)
        self.assertIn("LayoutGuidedDivider toolForm, M, y, contentW", helpers)
        self.assertIn("LayoutGuidedInfoBox toolForm, M, y, contentW", helpers)
        self.assertIn("LayoutGuidedPreviewButton toolForm, M, y, contentW", helpers)
        self.assertIn("LayoutGuidedScopeRow toolForm, M, y, contentW", helpers)
        self.assertIn("LayoutGuidedActionButtons toolForm, M, y, contentW", helpers)
        self.assertNotIn("Public Sub LayoutCapitalizationCleanupForm", helpers)
        self.assertNotIn("With toolForm.Controls(\"cmdHelp\")", helpers)
        self.assertNotIn("Case \"cmdPreview\", \"cmdRun\", \"cmdReset\", \"cmdHelp\"", helpers)
        self.assertNotIn("If Not ctl.Visible Then Exit Function", helpers)
        self.assertIn("If ShouldHideGuidedControl(toolForm, ctl.Name) Then Exit Function", helpers)
        self.assertIn("halfW = (contentW - GAP) / 2", helpers)

    def test_installer_generates_tool_forms_without_tool_help_buttons(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("Private Sub LayoutGenericToolControls(comp As VBIDE.VBComponent, designer As Object)", installer)
        self.assertIn('PositionControl designer, "cmdPreview", MARGIN, y, contentW, BTN_H', installer)
        self.assertIn('PositionControl designer, "cmdRun", MARGIN, y, halfW, BTN_H', installer)
        self.assertIn('PositionControl designer, "cmdReset", MARGIN + halfW + GAP, y, halfW, BTN_H', installer)
        self.assertIn("Private Sub RemoveToolFormHelpButton(formName As String, designer As Object)", installer)
        self.assertIn('designer.Controls.Remove "cmdHelp"', installer)
        self.assertNotIn('designer.Controls("cmdHelp").Visible = False', installer)
        self.assertNotIn('If CStr(cn) = "cmdRun" Or CStr(cn) = "cmdReset" Or CStr(cn) = "cmdHelp" Then', installer)
        self.assertNotIn('Case "cmdHelp": cap = "Help"', installer)

    def test_capitalization_still_preserves_behavior_choices(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('optAll.Caption = "Fix sentence starts"', form)
        self.assertIn('optSentence.Caption = "Sentence case"', form)
        self.assertIn('optTitle.Caption = "Title case"', form)
        self.assertIn('optUpper.Caption = "UPPERCASE"', form)
        self.assertIn('optLower.Caption = "lowercase"', form)
        self.assertIn('optCustom.Caption = "Custom"', form)
        self.assertIn("Private Sub UpdateModeSummary()", form)
        self.assertIn("Recommended: capitalizes likely sentence starts", form)
        self.assertIn("Custom runs selected choices in order", form)
        self.assertIn("LayoutCleanupToolForm Me", form)

        self.assertIn("If optAll.Value Then doSmart = True", form)
        self.assertIn("If optSentence.Value Then doSentence = True", form)
        self.assertIn("If optTitle.Value Then doTitle = True", form)
        self.assertIn("If optUpper.Value Then doUpper = True", form)
        self.assertIn("If optLower.Value Then doLower = True", form)
        self.assertIn(
            "If optCustom.Value Then doSentence = chkSentence.Value: doTitle = chkTitle.Value: doUpper = chkUpper.Value: doLower = chkLower.Value: doSmart = chkSmartSentences.Value",
            form,
        )


if __name__ == "__main__":
    unittest.main()
