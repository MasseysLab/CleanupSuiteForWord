from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class GuidedToolDynamicLayoutTests(unittest.TestCase):
    def test_shared_layout_hides_title_noise_and_inactive_controls(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("ApplyMSFormTitleStrategy toolForm, True", helpers)
        self.assertLess(
            helpers.index("ApplyMSFormTitleStrategy toolForm, True"),
            helpers.index('Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")'),
        )
        self.assertIn("Private Function ShouldHideGuidedControl(ByVal toolForm As Object, ByVal controlName As String) As Boolean", helpers)
        self.assertIn("If GuidedControlMatchesToolTitle(toolForm, ctl) Then", helpers)
        self.assertIn("Private Function GuidedControlMatchesToolTitle(ByVal toolForm As Object, ByVal ctl As Object) As Boolean", helpers)
        self.assertIn("Trim$(toolForm.Caption)", helpers)
        self.assertIn("If ShouldHideGuidedControl(toolForm, ctl.Name) Then", helpers)
        self.assertIn('HideGuidedControl toolForm.Controls("optScopeSelection")', helpers)
        self.assertIn('InStr(1, toolForm.Controls("optScopeDocument").Caption, "(always)", vbTextCompare) = 0', helpers)
        self.assertIn('And Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) > 0', helpers)

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

    def test_generic_guided_layout_pairs_choice_controls_in_two_columns(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("Dim pendingChoice As String", helpers)
        self.assertIn("FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP", helpers)
        self.assertIn("Private Sub FlushGuidedPendingChoice(ByVal toolForm As Object, ByRef pendingChoice As String", helpers)
        self.assertIn("Private Function GuidedSingleChoiceWidth(ByVal ctl As Object, ByVal contentW As Single) As Single", helpers)
        self.assertIn("Private Function GuidedChoiceRowHeight(ByVal firstControl As Object, Optional ByVal secondControl As Object = Nothing) As Single", helpers)
        self.assertIn("choiceW = (contentW - GAP) / 2", helpers)
        self.assertIn("If Len(Trim$(ctl.Caption)) > 42 Then", helpers)
        self.assertIn("GuidedSingleChoiceWidth = contentW - 24", helpers)
        self.assertIn("ElseIf Len(Trim$(ctl.Caption)) > 30 Then", helpers)
        self.assertIn("GuidedSingleChoiceWidth = pairW + 72", helpers)
        self.assertIn("toolForm.Controls(pendingChoice).Move M + ((contentW - singleChoiceW) / 2), y, singleChoiceW, rowH", helpers)
        self.assertIn("choiceCtl.Move M + choiceW + GAP + 8, y, choiceW - 8, rowH", helpers)
        self.assertNotIn("ctl.Move M + 8, y, contentW - 8, GuidedOptionHeight(ctl)", helpers)

        self.assertIn("Dim pendingChoice As String: pendingChoice = \"\"", installer)
        self.assertIn("FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP", installer)
        self.assertIn("Private Sub FlushGenericPendingChoice(designer As Object, ByRef pendingChoice As String", installer)
        self.assertIn("Private Function GenericSingleChoiceWidth(designer As Object, ByVal controlName As String, ByVal contentW As Single) As Single", installer)
        self.assertIn("choiceW = (contentW - GAP) / 2", installer)
        self.assertIn("If Len(Trim$(CStr(designer.Controls(controlName).Caption))) > 42 Then", installer)
        self.assertIn("GenericSingleChoiceWidth = contentW - 24", installer)
        self.assertIn("ElseIf Len(Trim$(CStr(designer.Controls(controlName).Caption))) > 30 Then", installer)
        self.assertIn("GenericSingleChoiceWidth = pairW + 72", installer)
        self.assertIn("PositionControl designer, pendingChoice, MARGIN + ((contentW - singleChoiceW) / 2), y, singleChoiceW, rowH", installer)
        self.assertIn("PositionControl designer, controlName, MARGIN + choiceW + GAP + 8, y, choiceW - 8, rowH", installer)
        self.assertNotIn("h = 18: lft = MARGIN + 8: w = contentW - 8", installer)

    def test_hidden_controls_do_not_force_guided_choice_row_breaks(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        hidden_block = helpers[
            helpers.index("If ShouldHideGuidedControl(toolForm, ctl.Name) Then"):
            helpers.index("If Left$(ctl.Name, 3) = \"fra\" Then")
        ]

        self.assertIn("If ShouldHideGuidedControl(toolForm, ctl.Name) Then", hidden_block)
        self.assertIn("HideGuidedControl ctl", hidden_block)
        self.assertNotIn("FlushGuidedPendingChoice toolForm, pendingChoice", hidden_block)
        self.assertNotIn("FlushGuidedPendingButton toolForm, pendingButton", hidden_block)

    def test_scope_controls_follow_preview_and_hide_when_not_supported(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        preview_actions = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Private Function ScopeSelectionIsAvailable(ByVal toolForm As Object) As Boolean", helpers)
        self.assertIn("Public Sub RefreshScopeSelectionState(ByVal toolForm As Object, Optional ByVal preferSelectionWhenAvailable As Boolean = False)", helpers)
        self.assertIn("selectionAvailable = (Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn)", helpers)
        self.assertIn("If preferSelectionWhenAvailable Then", helpers)
        self.assertIn('toolForm.Controls("optScopeSelection").Value = True', helpers)
        self.assertIn('toolForm.Controls("optScopeDocument").Value = True', helpers)
        self.assertIn("RefreshScopeSelectionState toolForm", helpers)
        self.assertIn('InStr(1, toolForm.Controls("optScopeDocument").Caption, "(always)", vbTextCompare) = 0', helpers)
        self.assertIn('And Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) > 0', helpers)
        self.assertIn('LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW', helpers)
        self.assertLess(
            helpers.index("LayoutGuidedInfoBox toolForm, M, y, contentW"),
            helpers.index("LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW"),
        )
        self.assertIn('HideGuidedControl toolForm.Controls("optScopeDocument")', helpers)
        self.assertIn('toolForm.Controls("optScopeSelection").Move scopeLeft, y, scopeW, 18', helpers)
        self.assertIn('toolForm.Controls("optScopeDocument").Move scopeLeft, y + 20, scopeW, 18', helpers)

        self.assertIn('LayoutGenericPreviewAndScopeRow designer, y, MARGIN, contentW, BTN_H', installer)
        self.assertLess(
            installer.index('FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP'),
            installer.index('LayoutGenericPreviewAndScopeRow designer, y, MARGIN, contentW, BTN_H'),
        )
        self.assertIn('InStr(1, CStr(designer.Controls("optScopeDocument").Caption), "(always)", vbTextCompare) = 0', installer)
        self.assertIn('And Len(Trim$(CStr(designer.Controls("optScopeSelection").Caption))) > 0', installer)
        self.assertIn('PositionControl designer, "optScopeSelection", scopeLeft, y, scopeW, 18', installer)
        self.assertIn('PositionControl designer, "optScopeDocument", scopeLeft, y + 20, scopeW, 18', installer)
        self.assertIn('PositionControl designer, "optScopeDocument", 0, 0, 0, 0', installer)
        self.assertIn('designer.Controls("lblTitle").Caption = ""', installer)
        self.assertIn('designer.Controls("lblTitle").Visible = False', installer)
        self.assertIn("LayoutCleanupToolFormForReconfigure src", preview_actions)
        self.assertNotIn("LayoutCleanupToolForm mSourceForm", preview_actions)

        selection_scoped_forms = [
            "src/forms/frmHyperlinkRemover.bas",
            "src/forms/frmUnicodeCleanup.bas",
            "src/forms/frmFormattingStripper.bas",
            "src/forms/frmTableCleaner.bas",
            "src/forms/frmFootnoteRemover.bas",
            "src/forms/frmFontNormalizer.bas",
            "src/forms/frmDuplicateDetector.bas",
            "src/forms/frmSpacingCleanup.bas",
            "src/forms/frmSoftReturnConverter.bas",
            "src/forms/frmCapitalizationCleanup.bas",
            "src/forms/frmPunctuationCleanup.bas",
            "src/forms/frmBreakNormalizer.bas",
            "src/forms/frmListCleanup.bas",
            "src/forms/frmParagraphCleanup.bas",
        ]
        for path in selection_scoped_forms:
            source = read(path)
            with self.subTest(form=path):
                self.assertIn("RefreshScopeSelectionState Me, True", source)

    def test_new_tool_guidance_records_two_column_rule(self):
        contributing = read("CONTRIBUTING.md")
        readme = read("README.md")
        mechanisms = read("docs/Cool_work_arounds_and_coding_mechanisms_for_VBA_programmed_through_Python.md")

        required_phrases = [
            "two-column guided layout",
            "normal opt/chk choices are paired into two columns",
            "Custom choices are still two-column choices",
            "single leftover choices are centered",
            "intro-first guided layout",
            "LayoutCleanupToolForm Me",
            "LayoutGenericToolControls",
        ]
        for phrase in required_phrases:
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, contributing)

        self.assertIn("two-column guided layout", readme)
        self.assertIn("single leftover choices are centered", readme)
        self.assertIn("two-column guided layout", mechanisms)
        self.assertIn("single leftover choices are centered", mechanisms)

    def test_custom_mode_forms_relayout_when_their_mode_changes(self):
        expected_modes = {
            "src/forms/frmUnicodeCleanup.bas": ["optAll", "optNBSP", "optZeroWidth", "optCustom"],
            "src/forms/frmPunctuationCleanup.bas": ["optAll", "optQuotes", "optDashes", "optEllipses", "optCustom"],
            "src/forms/frmSpacingCleanup.bas": ["optAll", "optDoubleSpaces", "optTrim", "optCustom"],
            "src/forms/frmListCleanup.bas": ["optAll", "optBullets", "optNumbering", "optIndent", "optCustom"],
            "src/forms/frmParagraphCleanup.bas": ["optAll", "optRemoveEmpty", "optNormalizeSpacing", "optCustom"],
            "src/forms/frmCapitalizationCleanup.bas": ["optAll", "optCustom"],
        }

        for path, controls in expected_modes.items():
            source = read(path)
            with self.subTest(form=path):
                if path.endswith("frmCapitalizationCleanup.bas"):
                    self.assertIn("UpdateAdvancedVisibility", source)
                    self.assertIn("If optCustom.Value Then", source)
                else:
                    self.assertIn("If optCustom.Value Then", source)
                self.assertNotIn("If fraCustom.Visible Then", source)
                for control in controls:
                    self.assertIn(f"Private Sub {control}_Click()", source)
                    self.assertIn("Layout", source)

    def test_select_all_buttons_only_show_for_active_custom_multiselects(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        capitalization = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn('Case "cmdSelectAll", "cmdDeselectAll"', helpers)
        self.assertIn("If Not GuidedCustomModeIsActive(toolForm) Then", helpers)
        self.assertIn("ShouldHideGuidedControl = True", helpers)
        self.assertIn("IsGuidedCustomChoiceControl(toolForm.Name, controlName)", helpers)
        self.assertIn('Case "chkSentence", "chkTitle", "chkUpper", "chkLower", "chkSmartSentences"', helpers)
        self.assertIn('Case "chkHeadingParentheses"', helpers)
        self.assertIn("LayoutCapitalizationChoices", helpers)
        self.assertIn("PositionCapitalizationGridRow", helpers)
        self.assertIn('ShouldHideGuidedControl = Not GuidedCustomModeIsActive(toolForm)', helpers)
        self.assertIn("optCustom.Value = True", capitalization)
        self.assertIn("cmdSelectAll.Visible = optCustom.Value", capitalization)
        self.assertIn("cmdDeselectAll.Visible = optCustom.Value", capitalization)

    def test_generated_guided_info_labels_are_reset_before_relayout(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("ResetGuidedGeneratedChrome toolForm", helpers)
        self.assertIn('RemoveGuidedGeneratedLabel toolForm, "lblGuidedToolName"', helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedDivider"', helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedInfoBorder"', helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedInfoLine" & CStr(i)', helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedInfoName" & CStr(i)', helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedInfoDetail" & CStr(i)', helpers)
        self.assertIn("Private Sub ResetGuidedGeneratedChrome(ByVal toolForm As Object)", helpers)
        self.assertIn("Private Sub RemoveGuidedGeneratedLabel(ByVal toolForm As Object, ByVal controlName As String)", helpers)
        self.assertIn("Private Sub HideGuidedGeneratedLabel(ByVal toolForm As Object, ByVal controlName As String)", helpers)

    def test_shared_layout_uses_compact_guided_spacing(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn(".Move M, 8, contentW - 52, introH", helpers)
        self.assertIn("y = 8 + introH + 4", helpers)
        self.assertIn('Case "lblGuidedIntro", "chkPreviewOnly", "chkHeadingParentheses", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"', helpers)
        self.assertIn("GuidedIntroHeight = 26", helpers)
        self.assertIn("GuidedIntroHeight = 16", helpers)
        self.assertIn("GuidedLabelHeight = 36", helpers)
        self.assertIn("GuidedLabelHeight = 24", helpers)
        self.assertIn("GuidedOptionHeight = 44", helpers)
        self.assertIn("GuidedOptionHeight = 28", helpers)
        self.assertIn("GuidedOptionHeight = 18", helpers)
        self.assertIn("Const GAP As Single = 4", helpers)
        self.assertIn("toolForm.Height = y + 26", helpers)

    def test_shared_layout_moves_explanations_into_the_guided_info_box(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        layout_block = helpers[
            helpers.index("Public Sub LayoutCleanupToolForm(ByVal toolForm As Object)"):
            helpers.index("Private Function EnsureGuidedLabel(ByVal toolForm As Object, ByVal controlName As String) As Object")
        ]

        self.assertIn("LayoutGuidedInfoBox toolForm, M, y, contentW", helpers)
        self.assertNotIn("LayoutGuidedWarningLabel toolForm, M, y, contentW", layout_block)
        self.assertIn("HideGuidedExternalWarnings toolForm", helpers)
        self.assertIn('HideGuidedControl toolForm.Controls("lblFuzzyWarning")', helpers)
        self.assertIn('HideGuidedControl toolForm.Controls("lblSpeedWarning")', helpers)
        self.assertIn('HideGuidedControl toolForm.Controls("lblModeSummary")', helpers)
        self.assertIn('Set nameCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoName" & CStr(i))', helpers)
        self.assertIn('Set detailCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoDetail" & CStr(i))', helpers)
        self.assertIn("Private Function GuidedInfoRecord(ByVal formName As String, ByVal controlName As String, ByVal fallbackText As String, ByVal isSelected As Boolean) As String", helpers)
        self.assertIn("Private Function GuidedInfoRecordName(ByVal rawLine As String) As String", helpers)
        self.assertIn("Private Function GuidedInfoRecordDetail(ByVal rawLine As String) As String", helpers)
        self.assertIn("Private Function GuidedInfoNameWidth(ByVal nameText As String) As Single", helpers)
        self.assertIn("Const BOX_PAD As Single = 10", helpers)
        self.assertIn("Const INNER_GAP As Single = 4", helpers)
        self.assertIn(".TextAlign = 1", helpers)
        self.assertIn(".Move M + BOX_PAD, lineTop, nameColW, lineH", helpers)
        self.assertIn(".Move M + BOX_PAD + nameColW + INNER_GAP, lineTop, detailW, lineH", helpers)
        self.assertIn("If nameW > nameColW Then nameColW = nameW", helpers)
        self.assertIn("If nameColW > 92 Then nameColW = 92", helpers)
        self.assertIn("detailW = contentW - ((BOX_PAD * 2) + nameColW + INNER_GAP)", helpers)
        self.assertIn('HideGuidedGeneratedLabel toolForm, "lblGuidedInfoMeasure"', helpers)
        self.assertIn('Set measureCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoMeasure")', helpers)
        self.assertIn("lineH = GuidedInfoLineHeight(toolForm, detailText, detailW, lineBold)", helpers)
        self.assertIn("lineTop = lineTop + lineH", helpers)
        self.assertIn("Private Function GuidedInfoLineHeight(ByVal toolForm As Object, ByVal detailText As String, ByVal detailWidth As Single, Optional ByVal isBold As Boolean = False) As Single", helpers)
        self.assertIn(".AutoSize = False", helpers)
        self.assertIn(".AutoSize = True", helpers)
        self.assertIn("GuidedInfoLineHeight = measureCtl.Height - 2", helpers)
        self.assertIn("If GuidedInfoLineHeight < 11 Then GuidedInfoLineHeight = 11", helpers)
        self.assertLess(
            helpers.index("HideGuidedExternalWarnings toolForm"),
            helpers.index("LayoutGuidedInfoBox toolForm, M, y, contentW"),
        )
        self.assertLess(
            helpers.index("LayoutGuidedInfoBox toolForm, M, y, contentW"),
            helpers.index("LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW"),
        )

    def test_guided_info_uses_tool_specific_names_and_advisories_for_shared_forms(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        expected_snippets = [
            'Case "frmUnicodeCleanup.optAll": GuidedInfoDisplayName = "All invisible characters"',
            'Case "frmUnicodeCleanup.optNBSP": GuidedInfoDisplayName = "Nonbreaking spaces"',
            'Case "frmUnicodeCleanup.optZeroWidth": GuidedInfoDisplayName = "Zero-width characters"',
            'Case "frmPunctuationCleanup.optAll": GuidedInfoDisplayName = "All punctuation fixes"',
            'Case "frmSpacingCleanup.optTrim": GuidedInfoDisplayName = "Trim outer spaces"',
            'Case "frmListCleanup.optBullets": GuidedInfoDisplayName = "Bullet conversion"',
            'Case "frmListCleanup.optNumbering": GuidedInfoDisplayName = "Number conversion"',
            'Case "frmListCleanup.chkHyphenToBullets": GuidedInfoDisplayName = "Hyphen lines"',
            'Case "frmUnicodeCleanup.optAll": GuidedInfoAdvisory = "Sweeps the common invisible troublemakers that interfere with search, wrap, and cleanup."',
            'Case "frmPunctuationCleanup.optQuotes": GuidedInfoAdvisory = "Targets curly quotes and apostrophes without touching dashes or ellipses."',
            'Case "frmSpacingCleanup.chkNormalizeAfterPunct": GuidedInfoAdvisory = "Resets uneven post-punctuation spacing to a single clean space."',
            'Case "frmListCleanup.chkNormalizeBullets": GuidedInfoAdvisory = "Turns manual bullets into the document\'s standard Word bullet form."',
            'GuidedGenericAdvisory = "Removes matching content wherever this cleanup finds it."',
            'GuidedGenericAdvisory = "Converts matching content into the target Word-friendly form."',
            'GuidedGenericAdvisory = "Applies this cleanup when the tool runs."',
        ]
        for snippet in expected_snippets:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, helpers)

    def test_non_custom_detail_groups_relayout_when_their_trigger_changes(self):
        break_form = read("src/forms/frmBreakNormalizer.bas")
        self.assertIn("Private Sub chkConvertSectionBreaks_Click()", break_form)
        self.assertIn("LayoutCleanupToolForm Me", break_form)

        duplicate_form = read("src/forms/frmDuplicateDetector.bas")
        for control in ["optMatchExact", "optMatchNormalized", "optMatchFuzzy"]:
            self.assertIn(f"Private Sub {control}_Click()", duplicate_form)
        self.assertIn("LayoutCleanupToolForm Me", duplicate_form)

    def test_duplicate_remover_uses_requested_two_column_match_layout(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        layout_start = helpers.index("Private Sub LayoutDuplicateChoices")
        layout_end = helpers.index("End Sub", layout_start)
        layout = helpers[layout_start:layout_end]

        self.assertIn('toolForm.Controls("optMatchExact").Move leftX, y', layout)
        self.assertIn('toolForm.Controls("optMatchFuzzy").Move rightX, y', layout)
        self.assertIn('toolForm.Controls("optMatchNormalized").Move leftX, y', layout)
        self.assertIn('toolForm.Controls("chkIncludeEmptyParagraphs").Move M + 8, y, contentW - 8', layout)
        self.assertIn('PositionDuplicateFuzzyThreshold toolForm, "optFuzzyLoose", rightX, y', layout)
        self.assertIn('PositionDuplicateFuzzyThreshold toolForm, "optFuzzyMedium", rightX, y', layout)
        self.assertIn('PositionDuplicateFuzzyThreshold toolForm, "optFuzzyStrict", rightX, y', layout)
        self.assertIn("LayoutGuidedDivider toolForm, M, y, contentW", layout)
        self.assertLess(layout.index('toolForm.Controls("optMatchExact").Move'), layout.index('toolForm.Controls("optMatchNormalized").Move'))
        self.assertLess(layout.index('toolForm.Controls("optMatchNormalized").Move'), layout.index('toolForm.Controls("chkIncludeEmptyParagraphs").Move'))
        self.assertLess(layout.index('PositionDuplicateFuzzyThreshold toolForm, "optFuzzyStrict"'), layout.index('toolForm.Controls("chkIncludeEmptyParagraphs").Move'))
        self.assertLess(layout.index("LayoutGuidedDivider toolForm, M, y, contentW"), layout.index('toolForm.Controls("chkIncludeEmptyParagraphs").Move'))

        header_form = read("src/forms/frmHeaderFooterStandardizer.bas")
        for control in ["optStandardize", "optClearAll", "chkAlignment"]:
            self.assertIn(f"Private Sub {control}_Click()", header_form)
        self.assertIn("LayoutCleanupToolForm Me", header_form)

    def test_launcher_global_settings_order_puts_review_first(self):
        installer = read("src/installer/installer.bas")
        controls_array = '"chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave", "lblAutoSaveWarning", "chkUpdateChecks", "cmdCheckUpdates", "cmdResetAll"'
        self.assertIn(controls_array, installer)
        self.assertLess(
            installer.index('PositionControl designer, "chkReturnToMainAfterApply"'),
            installer.index('PositionControl designer, "chkReturnToMainAfterClose"'),
        )

    def test_first_four_risk_forms_keep_accepted_reference_layouts(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        expected = {
            "frmPunctuationCleanup": "LayoutPunctuationRiskChoiceChips toolForm",
            "frmUnicodeCleanup": "LayoutUnicodeRiskChoiceChips toolForm",
            "frmSpacingCleanup": "LayoutSpacingRiskChoiceChips toolForm",
            "frmCapitalizationCleanup": "LayoutCapitalizationRiskChoiceChips toolForm",
        }

        for form_name, call in expected.items():
            with self.subTest(form=form_name):
                self.assertIn(f'Case "{form_name}"', helpers)
                self.assertIn(call, helpers)
        for path in [
            "src/forms/frmPunctuationCleanup.bas",
            "src/forms/frmUnicodeCleanup.bas",
            "src/forms/frmSpacingCleanup.bas",
            "src/forms/frmCapitalizationCleanup.bas",
        ]:
            with self.subTest(no_form_owned_layout=path):
                self.assertNotIn("RiskChoiceChips Me", read(path))

    def test_risk_placement_helpers_are_conformed_across_guided_tool_forms(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        expected_helpers = [
            "Public Sub LayoutUnicodeRiskChoiceChips",
            "Public Sub LayoutPunctuationRiskChoiceChips",
            "Public Sub LayoutSpacingRiskChoiceChips",
            "Public Sub LayoutCapitalizationRiskChoiceChips",
            "Public Sub LayoutListRiskChoiceChips",
            "Public Sub LayoutParagraphRiskChoiceChips",
            "Public Sub LayoutDuplicateRiskChoiceChips",
            "Public Sub LayoutFormattingRiskChoiceChips",
            "Public Sub LayoutFontRiskChoiceChips",
            "Public Sub LayoutStyleRiskChoiceChips",
            "Public Sub LayoutTableRiskChoiceChips",
            "Public Sub LayoutBreakRiskChoiceChips",
            "Public Sub LayoutHeaderFooterRiskChoiceChips",
            "Public Sub LayoutFootnoteRiskChoiceChips",
            "Public Sub LayoutObjectRiskChoiceChips",
            "Public Sub LayoutDocumentTrimRiskSummary",
            "Public Sub LayoutSoftReturnRiskChoiceChips",
            "Public Sub LayoutFinalReviewRiskSummary",
            "Public Sub LayoutHyperlinkRiskChoiceChips",
            "Private Sub PlaceRiskChipBesideChoice",
            "Private Function EnsureRiskPlacementChip",
            "Private Function EnsureRiskPlacementLabel",
            "Private Sub HideRiskPlacementControls",
        ]
        for snippet in expected_helpers:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, helpers)

        dispatcher_expectations = {
            "frmUnicodeCleanup": "LayoutUnicodeRiskChoiceChips toolForm",
            "frmPunctuationCleanup": "LayoutPunctuationRiskChoiceChips toolForm",
            "frmSpacingCleanup": "LayoutSpacingRiskChoiceChips toolForm",
            "frmCapitalizationCleanup": "LayoutCapitalizationRiskChoiceChips toolForm",
            "frmListCleanup": "LayoutListRiskChoiceChips toolForm",
            "frmParagraphCleanup": "LayoutParagraphRiskChoiceChips toolForm",
            "frmDuplicateDetector": "LayoutDuplicateRiskChoiceChips toolForm",
            "frmFormattingStripper": "LayoutFormattingRiskChoiceChips toolForm",
            "frmFontNormalizer": "LayoutFontRiskChoiceChips toolForm",
            "frmStyleCleanup": "LayoutStyleRiskChoiceChips toolForm",
            "frmTableCleaner": "LayoutTableRiskChoiceChips toolForm",
            "frmBreakNormalizer": "LayoutBreakRiskChoiceChips toolForm",
            "frmHeaderFooterStandardizer": "LayoutHeaderFooterRiskChoiceChips toolForm",
            "frmFootnoteRemover": "LayoutFootnoteRiskChoiceChips toolForm",
            "frmObjectRemover": "LayoutObjectRiskChoiceChips toolForm",
            "frmDocumentTrim": "LayoutDocumentTrimRiskSummary toolForm",
            "frmSoftReturnConverter": "LayoutSoftReturnRiskChoiceChips toolForm",
            "frmFinalReview": "LayoutFinalReviewRiskSummary toolForm",
            "frmHyperlinkRemover": "LayoutHyperlinkRiskChoiceChips toolForm",
        }
        for form_name, layout_call in dispatcher_expectations.items():
            with self.subTest(form=form_name):
                self.assertIn(f'Case "{form_name}"', helpers)
                self.assertIn(layout_call, helpers)

        self.assertNotIn("LayoutParagraphRiskColumnTest", helpers)
        self.assertNotIn("LayoutFinalReviewRiskSummaryTest", helpers)

        self.assertIn('If InStr(1, controlName, "RiskPlacement", vbTextCompare) > 0 Then', helpers)
        self.assertIn('"Current selection: Final action.', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementCapitalizationRecommended", "Text caution"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementUnicodeAll", "Safe cleanup"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementPunctuationAll", "Text caution"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementListAll", "Structure change"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "chkIncludeEmptyParagraphs", "cmdRiskPlacementDuplicateEmpty", "Removes content"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "chkConvertToText", "cmdRiskPlacementTableConvertToText", "Structure change"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "chkTables", "cmdRiskPlacementObjectTables", "Removes content"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optSoftToPara", "cmdRiskPlacementSoftReturnToPara", "Structure change"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optHideLinks"', helpers)
        self.assertIn('PlaceRiskChipBesideChoice toolForm, "optRemoveLinks"', helpers)

    def test_reconfigure_reapplies_tool_specific_risk_placement_layout(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        preview_actions = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Public Sub LayoutCleanupToolFormForReconfigure(ByVal toolForm As Object)", helpers)
        self.assertIn("LayoutCleanupToolForm toolForm", helpers)
        self.assertLess(
            helpers.index("toolForm.Height = y + 26", helpers.index("LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW")),
            helpers.index("ApplyGuidedRiskPlacementLayout toolForm", helpers.index("LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW")),
        )
        self.assertIn('Case "frmHyperlinkRemover"', helpers)
        self.assertIn("LayoutHyperlinkRiskChoiceChips toolForm", helpers)
        self.assertIn('Case "frmHeaderFooterStandardizer"', helpers)
        self.assertIn("LayoutHeaderFooterRiskChoiceChips toolForm", helpers)
        self.assertIn('Case "frmDocumentTrim"', helpers)
        self.assertIn("LayoutDocumentTrimRiskSummary toolForm", helpers)
        self.assertIn("LayoutCleanupToolFormForReconfigure src", preview_actions)

    def test_hyperlink_tool_uses_two_explicit_actions_with_clickable_risk_chips(self):
        form = read("src/forms/frmHyperlinkRemover.bas")
        installer = read("src/installer/installer.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('optHideLinks.Caption = "Hide hyperlink styling, keep links"', form)
        self.assertIn('optRemoveLinks.Caption = "Remove hyperlinks, keep visible text"', form)
        self.assertIn("optHideLinks.Value = True", form)
        self.assertIn("Private Sub cmdRiskPlacementHyperlinkHide_Click()", form)
        self.assertIn("Private Sub cmdRiskPlacementHyperlinkRemove_Click()", form)
        self.assertIn("ShowToolRiskChoiceExplanation", form)
        self.assertIn("ClearHyperlinkVisibleStyle hl.Range", form)
        self.assertIn("hl.Delete", form)
        self.assertNotIn("chkRemoveFormat", form)
        self.assertIn('Case "frmHyperlinkRemover.optHideLinks": ControlCaptionText = "Hide hyperlink styling, keep links"', installer)
        self.assertIn('Case "frmHyperlinkRemover.optRemoveLinks": ControlCaptionText = "Remove hyperlinks, keep visible text"', installer)
        self.assertIn('"optHideLinks", "optRemoveLinks", "cmdRiskPlacementHyperlinkHide", "cmdRiskPlacementHyperlinkRemove"', installer)
        self.assertNotIn('Case "frmHyperlinkRemover.chkRemoveFormat"', installer)
        self.assertIn('Case "frmHyperlinkRemover": GuidedToolIntro = "Choose whether to hide hyperlink styling or remove hyperlink targets."', helpers)

    def test_design_time_risk_buttons_have_click_handlers(self):
        installer = read("src/installer/installer.bas")
        forms_by_name = {
            "frmBreakNormalizer": "src/forms/frmBreakNormalizer.bas",
            "frmCapitalizationCleanup": "src/forms/frmCapitalizationCleanup.bas",
            "frmDuplicateDetector": "src/forms/frmDuplicateDetector.bas",
            "frmFontNormalizer": "src/forms/frmFontNormalizer.bas",
            "frmFootnoteRemover": "src/forms/frmFootnoteRemover.bas",
            "frmFormattingStripper": "src/forms/frmFormattingStripper.bas",
            "frmHeaderFooterStandardizer": "src/forms/frmHeaderFooterStandardizer.bas",
            "frmHyperlinkRemover": "src/forms/frmHyperlinkRemover.bas",
            "frmListCleanup": "src/forms/frmListCleanup.bas",
            "frmObjectRemover": "src/forms/frmObjectRemover.bas",
            "frmParagraphCleanup": "src/forms/frmParagraphCleanup.bas",
            "frmPunctuationCleanup": "src/forms/frmPunctuationCleanup.bas",
            "frmSoftReturnConverter": "src/forms/frmSoftReturnConverter.bas",
            "frmSpacingCleanup": "src/forms/frmSpacingCleanup.bas",
            "frmStyleCleanup": "src/forms/frmStyleCleanup.bas",
            "frmTableCleaner": "src/forms/frmTableCleaner.bas",
            "frmUnicodeCleanup": "src/forms/frmUnicodeCleanup.bas",
        }

        controls_function = installer[installer.index("Private Function ControlsForForm"):]
        for form_name, form_path in forms_by_name.items():
            form_case_start = controls_function.index(f'Case "{form_name}"')
            next_case = controls_function.find("\n        Case ", form_case_start + 1)
            form_block = controls_function[form_case_start: next_case if next_case != -1 else len(controls_function)]
            risk_controls = sorted(set(
                match.group(1)
                for match in __import__("re").finditer(r'"(cmdRiskPlacement[^"]+)"', form_block)
            ))
            source = read(form_path)
            for control_name in risk_controls:
                with self.subTest(form=form_name, control=control_name):
                    self.assertIn(f"Private Sub {control_name}_Click()", source)
                    self.assertIn("ShowToolRiskChoiceExplanation", source)

    def test_first_four_launcher_tools_use_hyperlink_reference_risk_chip_pattern(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        sync_script = read("scripts/sync_docm_code_only.ps1")

        self.assertIn('If InStr(1, CStr(cn), "RiskPlacement", vbTextCompare) > 0 Then', installer)
        for form_name, layout_name in {
            "frmUnicodeCleanup": "LayoutUnicodeRiskChoiceChips",
            "frmPunctuationCleanup": "LayoutPunctuationRiskChoiceChips",
            "frmSpacingCleanup": "LayoutSpacingRiskChoiceChips",
            "frmCapitalizationCleanup": "LayoutCapitalizationRiskChoiceChips",
        }.items():
            with self.subTest(form=form_name):
                self.assertIn(f'Case "{form_name}"', helpers)
                self.assertIn(f"{layout_name} toolForm", helpers)
                self.assertIn(f"Public Sub {layout_name}", helpers)

        self.assertNotIn("LayoutSpacingRiskInlineTest Me", read("src/forms/frmSpacingCleanup.bas"))
        self.assertNotIn("LayoutCapitalizationRiskGroupHeaderTest Me", read("src/forms/frmCapitalizationCleanup.bas"))
        self.assertNotIn("LayoutParagraphRiskColumnTest Me", read("src/forms/frmParagraphCleanup.bas"))
        self.assertNotIn("LayoutFinalReviewRiskSummaryTest Me", read("src/forms/frmFinalReview.bas"))

        expected_controls = [
            "cmdRiskPlacementUnicodeAll", "cmdRiskPlacementUnicodeNBSP", "cmdRiskPlacementUnicodeZeroWidth", "cmdRiskPlacementUnicodeCustom",
            "cmdRiskPlacementPunctuationAll", "cmdRiskPlacementPunctuationQuotes", "cmdRiskPlacementPunctuationDashes", "cmdRiskPlacementPunctuationEllipses", "cmdRiskPlacementPunctuationCustom",
            "cmdRiskPlacementSpacingAll", "cmdRiskPlacementSpacingDouble", "cmdRiskPlacementSpacingTrim", "cmdRiskPlacementSpacingCustom",
            "cmdRiskPlacementCapitalizationRecommended", "cmdRiskPlacementCapitalizationCustom",
            "cmdRiskPlacementListAll", "cmdRiskPlacementListBullets", "cmdRiskPlacementListNumbering", "cmdRiskPlacementListIndent", "cmdRiskPlacementListCustom",
            "cmdRiskPlacementParagraphAll", "cmdRiskPlacementParagraphRemoveEmpty", "cmdRiskPlacementParagraphSpacing", "cmdRiskPlacementParagraphCustom",
            "cmdRiskPlacementDuplicateExact", "cmdRiskPlacementDuplicateNormalized", "cmdRiskPlacementDuplicateFuzzy", "cmdRiskPlacementDuplicateEmpty",
            "cmdRiskPlacementFormattingChar", "cmdRiskPlacementFormattingPara", "cmdRiskPlacementFormattingQuick", "cmdRiskPlacementFormattingThorough", "cmdRiskPlacementFormattingStrip",
            "cmdRiskPlacementFontFace", "cmdRiskPlacementFontSize", "cmdRiskPlacementFontBold", "cmdRiskPlacementFontItalic", "cmdRiskPlacementFontColor",
            "cmdRiskPlacementStyleRemoveUnused", "cmdRiskPlacementStyleRemap",
            "cmdRiskPlacementTableRows", "cmdRiskPlacementTableCols", "cmdRiskPlacementTablePadding", "cmdRiskPlacementTableStripFormat", "cmdRiskPlacementTableBorders", "cmdRiskPlacementTableRemoveBorders", "cmdRiskPlacementTableConvertToText",
            "cmdRiskPlacementBreakSections", "cmdRiskPlacementBreakPages", "cmdRiskPlacementBreakConvert",
            "cmdRiskPlacementHeaderFooterStandardize", "cmdRiskPlacementHeaderFooterClear", "cmdRiskPlacementHeaderFooterBreakLinks",
            "cmdRiskPlacementFootnoteFootnotes", "cmdRiskPlacementFootnoteEndnotes", "cmdRiskPlacementFootnoteKeepText",
            "cmdRiskPlacementObjectPictures", "cmdRiskPlacementObjectTextBoxes", "cmdRiskPlacementObjectFrames", "cmdRiskPlacementObjectLines", "cmdRiskPlacementObjectControls", "cmdRiskPlacementObjectHiddenText", "cmdRiskPlacementObjectTables",
            "lblRiskPlacementDocumentTrimSummary",
            "cmdRiskPlacementSoftReturnToPara", "cmdRiskPlacementSoftReturnToSoft",
            "lblRiskPlacementFinalReviewSummary",
        ]
        for control_name in expected_controls:
            with self.subTest(control=control_name):
                self.assertIn(control_name, installer)
                self.assertIn(control_name, sync_script)

    def test_header_footer_layout_uses_two_column_order_and_singular_menu_text(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn('If toolForm.Name = "frmHeaderFooterStandardizer" Then', helpers)
        self.assertIn("LayoutHeaderFooterChoices toolForm, M, y, contentW, GAP", helpers)
        self.assertIn("Private Sub LayoutHeaderFooterChoices(ByVal toolForm As Object", helpers)
        self.assertIn('PositionGuidedChoicePairByName toolForm, "optStandardize", "optClearAll"', helpers)
        self.assertIn('PositionGuidedChoicePairByName toolForm, "chkHeaders", "chkFooters"', helpers)
        self.assertIn('PositionGuidedChoicePairByName toolForm, "chkFont", "chkSpacing"', helpers)
        self.assertIn('PositionGuidedChoicePairByName toolForm, "chkBreakLinks", "chkAlignment"', helpers)
        self.assertIn('PositionGuidedCenteredChoice toolForm, "optAlignCenter"', helpers)
        self.assertIn('PositionGuidedChoicePairByName toolForm, "optAlignLeft", "optAlignRight"', helpers)

        self.assertIn('Case "frmHeaderFooterStandardizer"', installer)
        self.assertIn('"chkHeaders", "chkFooters", "chkFont", "chkSpacing", "chkBreakLinks", "chkAlignment"', installer)
        self.assertLess(installer.index('"chkBreakLinks", "chkAlignment"'), installer.index('"fraAlign", "optAlignLeft"'))
        self.assertIn('"fraAlign", "optAlignLeft", "optAlignRight", "optAlignCenter"', installer)
        self.assertIn('Case "cmdHeaderFooter": cap = "Header / Footer"', installer)

    def test_header_footer_info_box_keeps_unlink_before_alignment(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('If toolForm.Name = "frmHeaderFooterStandardizer" Then', helpers)
        self.assertIn('GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkBreakLinks")', helpers)
        self.assertIn('GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkAlignment")', helpers)
        self.assertLess(
            helpers.index('GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkBreakLinks")'),
            helpers.index('GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkAlignment")'),
        )

    def test_capitalization_uses_shared_friendly_title_mapping(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('Case "frmCapitalizationCleanup": GuidedToolName = "Capitalization Fixer"', helpers)
        self.assertIn('Case "frmCapitalizationCleanup": GuidedToolIntro = "Use the recommended repair, or customize its protections."', helpers)


if __name__ == "__main__":
    unittest.main()
