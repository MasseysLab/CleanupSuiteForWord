from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class CapitalizationSmartRepairTests(unittest.TestCase):
    def test_form_offers_only_recommended_and_custom_repairs(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('lblIntro.Caption = "Use the recommended repair, or customize its protections."', form)
        self.assertIn('optAll.Caption = "Recommended repair"', form)
        self.assertIn('optCustom.Caption = "Custom repair"', form)
        self.assertIn('chkSentence.Caption = "Recognize common abbreviations"', form)
        self.assertIn('chkTitle.Caption = "Protect common acronyms"', form)
        self.assertIn('chkUpper.Caption = "Protect names and brands"', form)
        self.assertIn('chkLower.Caption = "Repair likely headings in title case"', form)
        self.assertIn('chkHeadingParentheses.Caption = "Also title-case words inside parentheses"', form)
        self.assertIn("chkHeadingParentheses.Value = False", form)
        self.assertIn('chkSmartSentences.Caption = "Recognize quotes, bullets, and sentence boundaries"', form)
        for obsolete in ("optSentence", "optTitle", "optUpper.GroupName", "optLower.GroupName", "coming soon", "Conservative", "Balanced (coming soon)", "Aggressive"):
            self.assertNotIn(obsolete, form)
        self.assertNotIn("fraCustom", form)

    def test_custom_controls_are_shown_only_for_custom_repair(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        for control in ("chkSentence", "chkTitle", "chkUpper", "chkLower", "chkSmartSentences", "cmdSelectAll", "cmdDeselectAll", "cmdEditExceptions"):
            self.assertIn(f"{control}.Visible = optCustom.Value", form)
        self.assertIn("chkHeadingParentheses.Visible = optCustom.Value And chkLower.Value", form)
        self.assertIn("chkHeadingParentheses.Enabled = optCustom.Value And chkLower.Value", form)
        self.assertIn("Private Sub ApplyModeDefaults()", form)
        self.assertIn("Private Function SmartRepairParagraph(", form)
        self.assertIn("Private Function ApplySentenceRepairs(", form)
        self.assertIn("Private Function NormalizeLongUppercaseWords(", form)
        self.assertIn("Private Function RestoreProtectedTerms(", form)
        self.assertIn("Private Function IsLikelyHeadingText(", form)

    def test_apply_path_uses_one_real_engine_without_dead_strength_modes(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn("If optAll.Value Then", form)
        self.assertIn("If optCustom.Value Then", form)
        self.assertIn("transformed = SmartRepairParagraph(", form)
        self.assertIn("Const LONG_UPPERCASE_THRESHOLD As Long = 8", form)
        self.assertNotIn("ActiveRepairModeName", form)
        self.assertNotIn("ShouldNormalizeEntireParagraph", form)
        self.assertNotIn("modeName As String", form)

    def test_numbered_headings_do_not_treat_the_number_as_the_first_title_word(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for snippet in [
            "Private Function CountHeadingWords(ByVal textValue As String) As Long",
            "Private Function TokenContainsLetter(ByVal tokenText As String) As Boolean",
            "IsFinalHeadingWord(textValue, i)",
            "If TokenContainsLetter(token) Then",
            "token = ApplyHeadingMajorWordCase(token)",
            "headingWordCount = CountHeadingWords(headingCandidate)",
            "CountHeadingCaseWords headingCandidate, titleWords, uppercaseWords",
            "titleWords >= headingWordCount - 1",
            "Public Function SmokeSmartRepairText(ByVal textValue As String) As String",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

        for snippet in [
            '"2. The Problem of Divine Hiddenness"',
            '"5. The Success of Evolutionary Biology"',
            '"9. The Argument from Nonbelief Across History"',
            '"10. The Burden of Proof / Lack of Empirical Evidence"',
            '"Capitalization Fixer changed a correctly capitalized numbered heading from ["',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, smoke)

        heading_case = form[form.index("Private Function ApplyHeadingCase("):form.index("Private Function IsHeadingMinorWord(")]
        self.assertNotIn("wordCount = CountWords(textValue)", heading_case)
        self.assertIn('"for", "from", "in"', form)

    def test_heading_repair_has_safe_parenthetical_subordinate_and_fail_closed_bypass(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        sync = read("scripts/sync_docm_code_only.ps1")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for snippet in [
            "headingLike = styleHeadingLike Or IsLikelyHeadingText(textValue)",
            "If headingLike And Not useHeadingHeuristics Then",
            "SmartRepairParagraph = textValue",
            "working = ApplyHeadingCase(working, titleCaseParentheticalText)",
            "RestoreOriginalParentheticalText(textValue, working)",
            "Private Function ApplyBalancedParentheticalHeadingCase(",
            "Private Function MatchingClosingParenthesis(",
            "Private Function TextOutsideParentheses(",
            "headingCandidate = Trim$(TextOutsideParentheses(trimmedText))",
            "Public Function SmokeSmartRepairWithoutHeadingsText(",
            "Public Function SmokeSmartRepairHeadingParenthesesText(",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

        for source in (helpers, installer, sync):
            self.assertIn("chkHeadingParentheses", source)
        self.assertIn("LayoutCapitalizationChoices", helpers)
        self.assertIn("PositionCapitalizationGridRow", helpers)
        self.assertIn('ControlCaptionText = "Also title-case words inside parentheses"', installer)

        for snippet in [
            '"Divine Hiddenness (weakened)"',
            '"Consciousness (the Hard Problem)"',
            '"Consciousness (The Hard Problem)"',
            '"Heading (outer words (inner words))"',
            '"Heading (Outer Words (Inner Words))"',
            '"Heading (first note) and (second note)"',
            '"Heading (First Note) and (Second Note)"',
            '"Kalam Cosmological Argument (a specific version of #1)"',
            '"Kalam Cosmological Argument (A Specific Version of #1)"',
            '"The Incoherence of God (Omnipotence, Omniscience, etc.)"',
            '"Teleological / Design in Nature (Biology, DNA)"',
            '"DNA carries genetic information."',
            '"Heading (unfinished words"',
            '"MOST RESILIENT ARGUMENTS FOR GOD"',
            '"This sentence has EXCESSIVELY loud text."',
            '"This sentence has Excessively loud text."',
            '"THE FINAL SHORTLIST: The Most Resilient Arguments Overall"',
            '"The Final Shortlist: The Most Resilient Arguments Overall"',
        ]:
            with self.subTest(smoke=snippet):
                self.assertIn(snippet, smoke)

    def test_dot_attached_domain_segments_stay_lowercase(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for snippet in [
            "Dim lowerDotSegment As Boolean",
            'If lowerDotSegment And currentChar Like "[A-Z]" Then currentChar = LCase$(currentChar)',
            "PeriodStartsLowercaseDotSegment(textValue, i)",
            "Private Function PeriodStartsLowercaseDotSegment(",
            'PeriodStartsLowercaseDotSegment = (segmentLength >= 2)',
            "If PeriodFollowsInitialChain(textValue, periodIndex) Then Exit Function",
            "Private Function PeriodFollowsInitialChain(",
            "Public Function SmokeSmartRepairHeadingText(",
            "If lowerDotSegment Then",
            "If Not startsLowerDotSegment Then",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

        for snippet in [
            '"Visit JW.Org, Example.COM, and Anything.BLAH."',
            '"Visit JW.org, Example.com, and Anything.blah."',
            '"Initials A.B., P.O.Box, and path Example.COM/MyPath stay readable."',
            '"Initials A.B., P.O.Box, and path Example.com/MyPath stay readable."',
            '"Resources at JW.Org and Example.COM"',
            '"Resources at JW.org and Example.com"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, smoke)

    def test_replaces_visible_paragraph_text_without_table_markers(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("original = ParagraphBodyText(p)", form)
        self.assertIn("ReplaceParagraphBodyText p, transformed", form)
        self.assertIn("Public Function ParagraphBodyText(ByVal paragraphItem As Paragraph) As String", helpers)
        self.assertIn("Public Sub ReplaceParagraphBodyText(ByVal paragraphItem As Paragraph, ByVal replacementText As String)", helpers)
        self.assertIn('Mid$(fullText, bodyLength, 1) = vbCr Or Mid$(fullText, bodyLength, 1) = Chr$(7)', helpers)
        self.assertIn("bodyRange.Text = replacementText", helpers)
        self.assertNotIn("p.Range.Text = transformed & vbCr", form)

    def test_custom_buttons_switch_into_custom_mode(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        select_all = form[form.index("Private Sub cmdSelectAll_Click()"):form.index("Private Sub cmdDeselectAll_Click()")]
        deselect_all = form[form.index("Private Sub cmdDeselectAll_Click()"):form.index("Private Sub chkSentence_Click()")]

        self.assertIn("optCustom.Value = True", select_all)
        self.assertIn("UpdateAdvancedVisibility", select_all)
        self.assertIn("optCustom.Value = True", deselect_all)
        self.assertIn("UpdateAdvancedVisibility", deselect_all)

    def test_guided_info_and_installer_match_the_two_choice_design(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        sync = read("scripts/sync_docm_code_only.ps1")

        self.assertIn('GuidedInfoDisplayName = "Recommended repair"', helpers)
        self.assertIn('GuidedInfoDisplayName = "Custom repair"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementCapitalizationRecommended", "Text caution"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "chkHeadingParentheses", "cmdRiskPlacementCapitalizationChkHeadingParentheses", "Text caution"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup.optAll": ControlCaptionText = "Recommended repair"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.optCustom": ControlCaptionText = "Custom repair"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.chkHeadingParentheses": ControlCaptionText = "Also title-case words inside parentheses"', installer)
        self.assertIn('"cmdRiskPlacementCapitalizationRecommended"', sync)
        for obsolete in ("Balanced (coming soon)", "Aggressive (coming soon)", "Legacy hidden option"):
            self.assertNotIn(obsolete, helpers + installer)
        for stale_control in ("optSentence", "optTitle", "optUpper", "optLower", "CapitalizationConservative", "CapitalizationBalanced", "CapitalizationAggressive"):
            self.assertNotIn(stale_control, form + helpers + installer + sync)

    def test_custom_choices_use_exactly_two_columns_and_three_rows(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for left_name, right_name in [
            ("chkSentence", "chkTitle"),
            ("chkUpper", "chkLower"),
            ("chkSmartSentences", "chkHeadingParentheses"),
        ]:
            with self.subTest(row=(left_name, right_name)):
                self.assertIn(
                    f'PositionCapitalizationGridRow toolForm, "{left_name}", "{right_name}"',
                    helpers,
                )
        self.assertIn('toolForm.Controls("chkSentence").Top = toolForm.Controls("chkTitle").Top', smoke)
        self.assertIn('toolForm.Controls("chkUpper").Top = toolForm.Controls("chkLower").Top', smoke)
        self.assertIn('toolForm.Controls("chkSmartSentences").Top = toolForm.Controls("chkHeadingParentheses").Top', smoke)

    def test_mixed_all_caps_and_title_case_line_is_treated_as_a_heading(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn("Private Sub CountHeadingCaseWords(", form)
        self.assertIn("uppercaseWords >= 2 And titleWords >= 2", form)
        self.assertIn('If currentChar = ":" Then startsHeadingSegment = True', form)
        self.assertIn('"THE FINAL SHORTLIST: The Most Resilient Arguments Overall"', smoke)
        self.assertIn('"The Final Shortlist: The Most Resilient Arguments Overall"', smoke)

    def test_parenthetical_minor_words_and_etc_stay_lowercase(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn('Like "[A-Za-z0-9]" Then Exit Function', form)
        self.assertIn("Private Function IsAlwaysLowercaseHeadingWord(", form)
        self.assertIn('LCase$(tokenText) = "etc"', form)
        self.assertIn('"Kalam Cosmological Argument (A Specific Version of #1)"', smoke)
        self.assertIn('"The Incoherence of God (Omnipotence, Omniscience, etc.)"', smoke)

    def test_straight_and_curly_possessive_suffixes_stay_lowercase(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn('currentChar = ChrW$(8217)', form)
        self.assertIn('"God\'s Attributes and Satan\'s Rulership"', smoke)
        self.assertIn('"God" & ChrW$(8217) & "s Attributes and JW.org"', smoke)

    def test_straight_and_word_smart_double_quotes_preserve_sentence_boundaries(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn('ChrW$(&H201C), ChrW$(&H201D)', form)
        self.assertIn('"""Quoted sentence."" Next sentence."', smoke)
        self.assertIn('ChrW$(&H201C) & "Quoted sentence."', smoke)

    def test_direct_quotes_after_a_comma_start_with_a_capital(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn("Private Function IsOpeningQuoteAfterComma(", form)
        self.assertIn('Case """", "\'", ChrW$(&H2018), ChrW$(&H201C)', form)
        self.assertIn('"He said,""Quoted words."""', smoke)
        self.assertIn('"He said,\'Quoted words.\'"', smoke)
        self.assertIn('ChrW$(&H201C) & "Quoted words."', smoke)
        self.assertIn('ChrW$(&H2018) & "Quoted words."', smoke)

    def test_dna_is_a_protected_canonical_acronym(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn('"DNA|DNA"', form)
        self.assertIn('"Teleological / Design in Nature (Biology, DNA)"', smoke)
        self.assertIn('"DNA carries genetic information."', smoke)

    def test_compiled_protection_lists_are_conservatively_extended(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for expected in [
            '"ave."',
            '"vols."',
            '"RNA|RNA"',
            '"MRNA|mRNA"',
            '"MRI|MRI"',
            '"XLSX|XLSX"',
            '"openai|OpenAI"',
            '"powershell|PowerShell"',
            '"macos|macOS"',
        ]:
            with self.subTest(expected=expected):
                self.assertIn(expected, form)

        acronym_lines = "\n".join(
            line for line in form.splitlines() if "mCapitalizationAcronymDict" in line
        )
        for unsafe_collision in ['"IT|IT"', '"US|US"', '"WHO|WHO"']:
            with self.subTest(unsafe_collision=unsafe_collision):
                self.assertNotIn(unsafe_collision, acronym_lines)

        self.assertIn('"DNA, RNA, mRNA, ATP, MRI, PDF, and XLSX remain protected."', smoke)
        self.assertIn('"OpenAI uses GitLab, PowerShell, and macOS."', smoke)
        self.assertIn('"Meet at Main Ave. near the park."', smoke)


if __name__ == "__main__":
    unittest.main()
