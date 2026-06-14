from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class GuidedToolDynamicLayoutTests(unittest.TestCase):
    def test_shared_layout_hides_stale_titles_and_inactive_controls(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Private Function ShouldHideGuidedControl(ByVal toolForm As Object, ByVal controlName As String) As Boolean", helpers)
        self.assertIn("If GuidedControlMatchesToolTitle(toolForm, ctl) Then", helpers)
        self.assertIn("Private Function GuidedControlMatchesToolTitle(ByVal toolForm As Object, ByVal ctl As Object) As Boolean", helpers)
        self.assertIn("If ShouldHideGuidedControl(toolForm, ctl.Name) Then", helpers)
        self.assertIn('HideGuidedControl toolForm.Controls("optScopeSelection")', helpers)
        self.assertIn('If Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) = 0 Then', helpers)

        conditional_names = [
            "chkNBSP", "chkZWSP", "chkZWNJ", "chkZWJ", "chkBOM", "chkSoftHyphen", "chkNBHyphen",
            "chkCurlyDouble", "chkCurlySingle", "chkEmDash", "chkEnDash", "chkEllipses",
            "chkDoubleSpaces", "chkTrimSpaces", "chkSpaceBeforePunct", "chkNormalizeAfterPunct", "chkExtraBlankLines",
            "chkNormalizeBullets", "chkNormalizeNumbering", "chkFixIndent", "chkHyphenToBullets",
            "chkRemoveEmpty", "chkCollapseBreaks", "chkNormalizeParaSpacing",
            "optConvertNextPage", "optConvertContinuous", "optConvertEvenPage", "optConvertOddPage",
            "optFuzzyLoose", "optFuzzyMedium", "optFuzzyStrict",
            "optAlignLeft", "optAlignCenter", "optAlignRight",
        ]
        for name in conditional_names:
            with self.subTest(control=name):
                self.assertIn(f'"{name}"', helpers)

    def test_generic_guided_tools_use_two_column_choice_rows(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("Dim pendingChoice As String", helpers)
        self.assertIn("FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP", helpers)
        self.assertIn("Private Sub FlushGuidedPendingChoice(ByVal toolForm As Object, ByRef pendingChoice As String", helpers)
        self.assertIn("Private Function GuidedChoiceRowHeight(ByVal firstControl As Object, Optional ByVal secondControl As Object = Nothing) As Single", helpers)
        self.assertIn("choiceW = (contentW - GAP) / 2", helpers)
        self.assertIn("toolForm.Controls(pendingChoice).Move M + 8, y, choiceW - 8, rowH", helpers)
        self.assertIn("choiceCtl.Move M + choiceW + GAP + 8, y, choiceW - 8, rowH", helpers)
        self.assertNotIn("ctl.Move M + 8, y, contentW - 8, GuidedOptionHeight(ctl)", helpers)

        self.assertIn("Dim pendingChoice As String: pendingChoice = \"\"", installer)
        self.assertIn("FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP", installer)
        self.assertIn("Private Sub FlushGenericPendingChoice(designer As Object, ByRef pendingChoice As String", installer)
        self.assertIn("choiceW = (contentW - GAP) / 2", installer)
        self.assertIn("PositionControl designer, pendingChoice, MARGIN + 8, y, choiceW - 8, rowH", installer)
        self.assertIn("PositionControl designer, controlName, MARGIN + choiceW + GAP + 8, y, choiceW - 8, rowH", installer)
        self.assertNotIn("h = 18: lft = MARGIN + 8: w = contentW - 8", installer)

    def test_new_tool_guidance_records_two_column_rule(self):
        contributing = read("CONTRIBUTING.md")
        readme = read("README.md")
        mechanisms = read("docs/Cool_work_arounds_and_coding_mechanisms_for_VBA_programmed_through_Python.md")

        required_phrases = [
            "two-column guided layout",
            "normal opt/chk choices are paired into two columns",
            "Custom choices are still two-column choices",
            "LayoutCleanupToolForm Me",
            "LayoutGenericToolControls",
        ]
        for phrase in required_phrases:
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, contributing)

        self.assertIn("two-column guided layout", readme)
        self.assertIn("two-column guided layout", mechanisms)

    def test_custom_forms_relayout_when_mode_changes(self):
        expected_modes = {
            "src/forms/frmUnicodeCleanup.bas": ["optAll", "optNBSP", "optZeroWidth", "optCustom"],
            "src/forms/frmPunctuationCleanup.bas": ["optAll", "optQuotes", "optDashes", "optEllipses", "optCustom"],
            "src/forms/frmSpacingCleanup.bas": ["optAll", "optDoubleSpaces", "optTrim", "optCustom"],
            "src/forms/frmListCleanup.bas": ["optAll", "optBullets", "optNumbering", "optIndent", "optCustom"],
            "src/forms/frmParagraphCleanup.bas": ["optAll", "optRemoveEmpty", "optNormalizeSpacing", "optCustom"],
            "src/forms/frmCapitalizationCleanup.bas": ["optAll", "optSentence", "optTitle", "optUpper", "optLower", "optCustom"],
        }

        for path, controls in expected_modes.items():
            source = read(path)
            with self.subTest(form=path):
                self.assertIn("If optCustom.Value Then", source)
                self.assertNotIn("If fraCustom.Visible Then", source)
                for control in controls:
                    self.assertIn(f"Private Sub {control}_Click()", source)
                    self.assertIn("Layout", source)

    def test_non_custom_detail_groups_relayout_when_trigger_changes(self):
        break_form = read("src/forms/frmBreakNormalizer.bas")
        self.assertIn("Private Sub chkConvertSectionBreaks_Click()", break_form)
        self.assertIn("LayoutCleanupToolForm Me", break_form)

        duplicate_form = read("src/forms/frmDuplicateDetector.bas")
        for control in ["optMatchExact", "optMatchNormalized", "optMatchFuzzy"]:
            self.assertIn(f"Private Sub {control}_Click()", duplicate_form)
        self.assertIn("LayoutCleanupToolForm Me", duplicate_form)

        header_form = read("src/forms/frmHeaderFooterStandardizer.bas")
        for control in ["optStandardize", "optClearAll", "chkAlignment"]:
            self.assertIn(f"Private Sub {control}_Click()", header_form)
        self.assertIn("LayoutCleanupToolForm Me", header_form)

    def test_launcher_global_settings_order_puts_review_first(self):
        installer = read("src/installer/installer.bas")
        controls_array = '"chkShowCompletionReviewAfterApply", "chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave", "cmdResetAll"'
        self.assertIn(controls_array, installer)
        self.assertLess(
            installer.index('PositionControl designer, "chkShowCompletionReviewAfterApply"'),
            installer.index('PositionControl designer, "chkReturnToMainAfterApply"'),
        )


if __name__ == "__main__":
    unittest.main()
