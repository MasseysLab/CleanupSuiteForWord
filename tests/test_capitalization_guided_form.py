from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class CapitalizationSmartRepairTests(unittest.TestCase):
    def test_capitalization_form_uses_conservative_plus_custom_and_disabled_coming_soon_modes(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('lblIntro.Caption = "Choose how strongly capitalization should be repaired."', form)
        self.assertIn('optAll.Caption = "Conservative"', form)
        self.assertIn('optSentence.Caption = "Balanced (coming soon)"', form)
        self.assertIn('optTitle.Caption = "Aggressive (coming soon)"', form)
        self.assertIn('optSentence.Enabled = False', form)
        self.assertIn('optTitle.Enabled = False', form)
        self.assertIn('optUpper.Visible = False', form)
        self.assertIn('optLower.Visible = False', form)
        self.assertIn('optCustom.Caption = "Custom"', form)
        self.assertIn('optCustom.Visible = True', form)
        self.assertIn('chkSentence.Caption = "Use abbreviation lists"', form)
        self.assertIn('chkTitle.Caption = "Protect common acronyms"', form)
        self.assertIn('chkUpper.Caption = "Protect names and brands"', form)
        self.assertIn('chkLower.Caption = "Treat likely headings carefully"', form)
        self.assertIn('chkSmartSentences.Caption = "Use quote, bullet, and boundary context"', form)
        self.assertIn("cmdSelectAll.Visible = optCustom.Value", form)
        self.assertIn("cmdDeselectAll.Visible = optCustom.Value", form)
        self.assertIn("chkSentence.Visible = optCustom.Value", form)
        self.assertIn("chkTitle.Visible = optCustom.Value", form)
        self.assertIn("chkUpper.Visible = optCustom.Value", form)
        self.assertIn("chkLower.Visible = optCustom.Value", form)
        self.assertIn("chkSmartSentences.Visible = optCustom.Value", form)
        self.assertIn("Private Sub ApplyModeDefaults()", form)
        self.assertIn("Private Function ActiveRepairModeName(Optional ByVal unused As Boolean = False) As String", form)
        self.assertIn("Private Function SmartRepairParagraph(", form)
        self.assertIn("Private Function ApplySentenceRepairs(", form)
        self.assertIn("Private Function NormalizeLongUppercaseWords(", form)
        self.assertIn("Private Function RestoreProtectedTerms(", form)
        self.assertIn("Private Function IsLikelyHeadingText(", form)
        self.assertIn("Private Function GetCapitalizationAbbreviationDict() As Object", form)
        self.assertIn("Private Function GetCapitalizationAcronymDict() As Object", form)
        self.assertIn("Private Function GetCapitalizationProtectedTermDict() As Object", form)

    def test_capitalization_apply_path_uses_smart_repair_engine_instead_of_word_duplicate_modes(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('If optAll.Value Then', form)
        self.assertIn('ActiveRepairModeName = "Conservative"', form)
        self.assertIn('If optCustom.Value Then', form)
        self.assertIn('ActiveRepairModeName = "Custom"', form)
        self.assertIn("transformed = SmartRepairParagraph(", form)
        self.assertNotIn("transformed = SentenceCase(transformed)", form)
        self.assertNotIn("transformed = TitleCase(transformed)", form)
        self.assertNotIn("transformed = UCase$(transformed)", form)
        self.assertNotIn("transformed = LCase$(transformed)", form)
        self.assertNotIn("doSentence As Boolean", form)
        self.assertNotIn("doTitle As Boolean", form)
        self.assertNotIn("doUpper As Boolean", form)
        self.assertNotIn("doLower As Boolean", form)

    def test_custom_buttons_switch_capitalization_form_into_custom_mode(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        select_all_block = form[
            form.index("Private Sub cmdSelectAll_Click()"):form.index("Private Sub cmdDeselectAll_Click()")
        ]
        deselect_all_block = form[
            form.index("Private Sub cmdDeselectAll_Click()"):form.index("Private Sub chkSentence_Click()")
        ]

        self.assertIn("optCustom.Value = True", select_all_block)
        self.assertIn("UpdateAdvancedVisibility", select_all_block)
        self.assertIn("optCustom.Value = True", deselect_all_block)
        self.assertIn("UpdateAdvancedVisibility", deselect_all_block)

    def test_guided_info_and_help_text_describe_smart_repair_tool(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('Case "frmCapitalizationCleanup.optAll": GuidedInfoDisplayName = "Conservative repair"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.optSentence": GuidedInfoDisplayName = "Balanced (coming soon)"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.optTitle": GuidedInfoDisplayName = "Aggressive (coming soon)"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.optCustom": GuidedInfoDisplayName = "Custom"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.chkSentence": GuidedInfoDisplayName = "Abbreviation lists"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.chkTitle": GuidedInfoDisplayName = "Acronym protection"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.chkUpper": GuidedInfoDisplayName = "Name and brand protection"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.chkLower": GuidedInfoDisplayName = "Heading heuristics"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.chkSmartSentences": GuidedInfoDisplayName = "Boundary context"', helpers)
        self.assertIn("Fixes only obvious capitalization damage", helpers)
        self.assertIn("Planned for a later version", helpers)
        self.assertIn("Capitalization Fixer", helpers)
        self.assertIn("Conservative repair", helpers)
        self.assertIn("Balanced (coming soon)", helpers)
        self.assertIn("Aggressive (coming soon)", helpers)
        self.assertIn("Custom", helpers)
        self.assertNotIn("Title case: capitalizes the first letter of every word.", helpers)
        self.assertNotIn("UPPERCASE / lowercase: converts affected paragraph text to that case.", helpers)
        self.assertIn('Case "chkSentence", "chkTitle", "chkUpper", "chkLower", "chkSmartSentences"', helpers)
        self.assertIn('ShouldHideGuidedControl = Not GuidedCustomModeIsActive(toolForm)', helpers)

    def test_installer_and_layout_support_capitalization_advanced_section(self):
        installer = read("src/installer/installer.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('Case "frmCapitalizationCleanup.optAll": ControlCaptionText = "Conservative"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optSentence": ControlCaptionText = "Balanced (coming soon)"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optTitle": ControlCaptionText = "Aggressive (coming soon)"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optCustom": ControlCaptionText = "Custom"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkSentence": ControlCaptionText = "Use abbreviation lists"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkTitle": ControlCaptionText = "Protect common acronyms"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkUpper": ControlCaptionText = "Protect names and brands"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkLower": ControlCaptionText = "Treat likely headings carefully"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkSmartSentences": ControlCaptionText = "Use quote, bullet, and boundary context"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.lblIntro": ControlCaptionText = "Choose how strongly capitalization should be repaired."', installer)
        self.assertIn("Private Function IsCapitalizationAdvancedControl", helpers)
        self.assertIn('toolForm.Name = "frmCapitalizationCleanup"', helpers)
        self.assertIn("IsCapitalizationAdvancedControl(ctl.Name)", helpers)
        self.assertNotIn('Case "frmCapitalizationCleanup"' + "\n            IsGuidedCustomChoiceControl =", helpers)

    def test_future_work_notes_record_later_versions(self):
        roadmap = read("docs/Capitalization_Smart_Repair_Followups.md")

        self.assertIn("custom exception list", roadmap)
        self.assertIn("larger compiled data packs", roadmap)
        self.assertIn("style-aware heading detection", roadmap)
        self.assertIn("document-level caching", roadmap)


if __name__ == "__main__":
    unittest.main()
