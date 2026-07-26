from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def procedure_blocks(source: str):
    pattern = re.compile(
        r"(?im)^\s*(?:Private|Public)?\s*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)[^\r\n]*"
    )
    matches = list(pattern.finditer(source))
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(source)
        yield match.group(1), source[match.start():end]


def dim_names(line: str):
    stripped = line.split("'", 1)[0].strip()
    if not stripped.lower().startswith("dim "):
        return []
    declaration = stripped[4:]
    candidate = declaration.split(",", 1)[0].strip()
    names = [re.split(r"\s+As\s+|\s*:", candidate, maxsplit=1, flags=re.IGNORECASE)[0].strip()]
    return [name.lower() for name in names if name]


class ToolCompileRegressionTests(unittest.TestCase):
    def test_formatting_cleaner_uses_supported_paragraph_reset_api(self):
        source = read("src/forms/frmFormattingStripper.bas")

        self.assertNotIn("ClearParagraphDirectFormatting", source)
        self.assertIn("p.Range.ParagraphFormat.Reset", source)

    def test_break_normalizer_deletes_repeated_break_ranges_without_replace_tokens(self):
        source = read("src/forms/frmBreakNormalizer.bas")

        self.assertIn('cntSect = CollapseRepeatedBreaks(targetRange, "^b^b")', source)
        self.assertIn('cntPage = CollapseRepeatedBreaks(targetRange, "^m^m")', source)
        self.assertIn("deleteRange.Delete", source)
        self.assertNotIn('.Replacement.Text = "^b"', source)

    def test_style_cleanup_remaps_used_paragraphs_without_find_replace_or_style_deletion(self):
        source = read("src/forms/frmStyleCleanup.bas")
        remap_start = source.index("Private Function RemapStyle")
        remap_end = source.index("Private Sub chkRemoveUnused_Click", remap_start)
        remap_block = source[remap_start:remap_end]

        self.assertIn("For Each paragraphItem In storyPart.Paragraphs", remap_block)
        self.assertIn("paragraphItem.Range.Style = targetStyle", remap_block)
        self.assertNotIn("wdReplaceAll", remap_block)
        self.assertNotIn(".Delete", remap_block)

    def test_style_cleanup_preview_counts_used_variant_paragraphs(self):
        source = read("src/forms/frmStyleCleanup.bas")
        count_start = source.index("Private Function CountStyleVariantsForRemap")
        count_block = source[count_start:]

        self.assertIn("CountParagraphsUsingStyle(variantStyle)", count_block)
        self.assertNotIn("CountStyleVariantsForRemap = CountStyleVariantsForRemap + 1", count_block)

    def test_cleanup_start_does_not_save_when_global_auto_save_is_off(self):
        source = read("src/modules/modCleanupHelpers.bas")
        start = source.index("Public Sub MarkCleanupStart")
        end = source.index("Public Sub MarkCleanupEnd", start)

        self.assertNotIn("ActiveDocument.Save", source[start:end])

    def test_object_remover_counts_content_controls_as_form_controls(self):
        source = read("src/forms/frmObjectRemover.bas")

        self.assertIn("ActiveDocument.ContentControls.Count", source)
        self.assertIn("ActiveDocument.ContentControls(k).Delete True", source)

    def test_unicode_preview_has_visibility_aware_adjacent_fallbacks(self):
        source = read("src/forms/frmUnicodeCleanup.bas")

        self.assertIn("ActiveWindow.View.ShowAll", source)
        self.assertIn("HighlightUnicodeAdjacentCharacter", source)
        self.assertIn("ApplyPreviewHighlight markerRange", source)

    def test_capitalization_preview_highlights_changed_characters_and_repairs_selected_protections(self):
        source = read("src/forms/frmCapitalizationCleanup.bas")

        self.assertIn("HighlightCapitalizationDifferenceRuns p, original, transformed", source)
        self.assertIn("Do While i <= compareLength", source)
        self.assertIn("working = ApplyHeadingCase(working, titleCaseParentheticalText)", source)
        self.assertIn('Case "dr": CanonicalProtectedToken = "Dr"', source)

    def test_paragraph_cleanup_previews_and_removes_manual_leading_indents(self):
        source = read("src/forms/frmParagraphCleanup.bas")

        self.assertIn("ParagraphLeadingIndentLength(p)", source)
        self.assertIn("ApplyPreviewHighlight ActiveDocument.Range(p.Range.Start, p.Range.Start + leadingIndentLength)", source)
        self.assertIn("ActiveDocument.Range(p.Range.Start, p.Range.Start + leadingIndentLength).Delete", source)

    def test_spacing_tool_has_no_duplicate_local_dim_names(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        for proc_name, block in procedure_blocks(source):
            seen = set()
            duplicates = []
            for line in block.splitlines():
                for name in dim_names(line):
                    if name in seen:
                        duplicates.append(name)
                    seen.add(name)
            self.assertEqual([], duplicates, f"{proc_name} has duplicate local declarations")

    def test_table_cleaner_uses_valid_word_range_formatting_calls(self):
        source = read("src/forms/frmTableCleaner.bas")

        self.assertNotIn("tbl.Range.ClearFormatting", source)
        self.assertIn("tbl.Range.Font.Reset", source)
        self.assertIn("tbl.Range.ParagraphFormat.Reset", source)

    def test_spacing_tool_does_not_use_invalid_word_wildcard_punctuation_pattern(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        self.assertNotIn('" ([.,;:!?])"', source)
        self.assertNotIn(".Replacement.Text = \"\\1\"", source)
        self.assertIn("beforePunctPatterns = Array(\" .\", \" ,\", \" ;\", \" :\", \" !\", \" ?\")", source)
        self.assertIn("previewBeforePunctPatterns = Array(\" .\", \" ,\", \" ;\", \" :\", \" !\", \" ?\")", source)

    def test_metadatasuite_sharing_properties_preview_has_no_blank_line_after_heading(self):
        source = read("src/modules/modMetaDataSuite.bas")

        self.assertIn('"Clear Sharing Properties preview:" & vbCrLf & vbCrLf', source)
        self.assertNotIn('"Clear Sharing Properties preview:" & vbCrLf & vbCrLf & vbCrLf', source)

    def test_punctuation_tool_uses_bulk_replacement_with_smart_quotes_disabled(self):
        source = read("src/forms/frmPunctuationCleanup.bas")

        self.assertIn("Private Function ReplacePunctuationCharacterSet(", source)
        self.assertIn("Options.AutoFormatAsYouTypeReplaceQuotes = False", source)
        self.assertIn("Options.AutoFormatReplaceQuotes = False", source)
        self.assertIn("ReplacePunctuationCharacterSet = ReplacePunctuationCharacterSet + CountPreviewFindMatches", source)
        self.assertIn('.Text = "[" & findCharacters & "]"', source)
        self.assertIn(".MatchWildcards = True", source)
        self.assertIn(".Execute Replace:=wdReplaceAll", source)
        self.assertNotIn("matchRange.Text = replacementText", source)

    def test_punctuation_preview_uses_bulk_formatted_replace_all(self):
        source = read("src/forms/frmPunctuationCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        preview_start = source.index("If previewOnly Then")
        preview_end = source.index('MarkCleanupStart "Punctuation Normalizer"', preview_start)
        preview_block = source[preview_start:preview_end]
        helper_start = helpers.index("Public Sub HighlightPreviewFindCharacters")
        helper_end = helpers.index("End Sub", helper_start)
        highlight_helper = helpers[helper_start:helper_end]

        self.assertIn("HighlightPreviewFindCharacters targetRange, previewCharacters", preview_block)
        self.assertNotIn(".Replacement.Highlight = True", preview_block)
        self.assertNotIn("Replace:=wdReplaceAll", preview_block)
        self.assertIn("highlightRange.Find.HitHighlight", highlight_helper)
        self.assertIn('FindText:="[" & findCharacters & "]"', highlight_helper)
        self.assertIn("HighlightColor:=highlightColor", highlight_helper)
        self.assertIn("MatchWildcards:=True", highlight_helper)
        self.assertNotIn(".Execute Replace:=wdReplaceAll", highlight_helper)
        self.assertIn("highlightColor As WdColor = wdColorYellow", highlight_helper)
        self.assertNotIn("Application.ScreenUpdating = False", preview_block)

    def test_paragraph_spacing_preview_marks_boundaries_without_shading_text(self):
        source = read("src/forms/frmParagraphCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        preview_start = source.index("If previewOnly Then")
        preview_end = source.index('MarkCleanupStart "Paragraph Fixer"', preview_start)
        preview_block = source[preview_start:preview_end]
        spacing_start = preview_block.index("If doNormalize Then")
        spacing_end = preview_block.index("If doIndent Then", spacing_start)
        spacing_block = preview_block[spacing_start:spacing_end]

        self.assertIn("ApplyPreviewSpacingBoundary p.Range", spacing_block)
        self.assertNotIn("ApplyPreviewShading p.Range", spacing_block)
        self.assertIn("Public Sub ApplyPreviewSpacingBoundary", helpers)
        self.assertIn("markerColor As WdColorIndex = wdBrightGreen", helpers)
        self.assertIn("trailingCharacter <> vbCr And trailingCharacter <> Chr$(7)", helpers)
        self.assertIn("ApplyPreviewNearestVisibleCharacter(paragraphRange, markerColor)", helpers)
        self.assertIn("Private Function ApplyPreviewNearestVisibleCharacter", helpers)
        self.assertIn("paragraphRange.Document.Range(paragraphRange.Start - 1, paragraphRange.Start)", helpers)
        self.assertIn("If trailingCharacter = Chr$(7) Then Exit Function", helpers)
        self.assertIn("nearbyRange.Start = nearbyRange.Start - 1", helpers)
        self.assertIn("ApplyPreviewShading paragraphRange", helpers)
        self.assertIn("markerRange.Start = markerRange.End - 1", helpers)
        self.assertIn("ApplyPreviewHighlight markerRange, markerColor", helpers)
        self.assertNotIn("RestorePreviewSpacingBoundaries", helpers)

    def test_unicode_tool_applies_revalidated_engine_candidates_bottom_up(self):
        source = read("src/forms/frmUnicodeCleanup.bas")

        self.assertIn(
            "HybridRevalidateInvisibleUnicode(mHybridAnalysis, targetRange, applyFailure)",
            source,
        )
        self.assertIn(
            "For idx = applyCandidates.Count To 1 Step -1",
            source,
        )
        self.assertIn(
            "Set applyRange = HybridCandidateAbsoluteRange",
            source,
        )
        self.assertIn(
            "applyRange.Text = HybridCandidateReplacement(applyCandidate)",
            source,
        )
        self.assertNotIn("While .Execute(Replace:=wdReplaceOne)", source)

    def test_unicode_preview_summary_treats_invisible_characters_as_highlightable(self):
        source = read("src/forms/frmUnicodeCleanup.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("ShowPreviewActionsSummary Me, \"Unicode Cleaner\", previewRows")]

        self.assertIn('AddPreviewSummaryRow previewRows, "Non-breaking spaces", previewNBSP, False, (Not includeNBSP)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Zero-width spaces", previewZWSP, False, (Not includeZWSP)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Zero-width non-joiners", previewZWNJ, False, (Not includeZWNJ)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Zero-width joiners", previewZWJ, False, (Not includeZWJ)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Byte order marks", previewBOM, False, (Not includeBOM)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Soft hyphens", previewSoftHyphen, False, (Not includeSoftHyphen)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Non-breaking hyphens", previewNBHyphen, False, (Not includeNBHyphen)', preview_block)

    def test_spacing_tool_counts_double_space_replacements_without_mopup_gap(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        self.assertIn("Do", source)
        self.assertIn("replacedThisPass = ReplaceSpacingLiteralMatches(targetRange, \"  \", \" \")", source)
        self.assertIn("Loop While replacedThisPass > 0", source)
        self.assertIn("Private Function ReplaceSpacingLiteralMatches(", source)
        self.assertIn("matchRange.Text = replacementText", source)
        self.assertNotIn("Loop While targetRange.Find.Execute(FindText:=\"  \", Forward:=True)", source)

    def test_spacing_trim_replaces_visible_paragraph_text_without_table_markers(self):
        source = read("src/forms/frmSpacingCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("paraText = ParagraphBodyText(para)", source)
        self.assertIn("ReplaceParagraphBodyText para, trimmedText", source)
        self.assertNotIn("SpacingParagraphVisibleText", source)
        self.assertNotIn("ReplaceSpacingParagraphVisibleText", source)
        self.assertIn('Mid$(fullText, bodyLength, 1) = vbCr Or Mid$(fullText, bodyLength, 1) = Chr$(7)', helpers)
        self.assertNotIn("para.Range.Text = trimmedText & vbCr", source)

    def test_spacing_tool_reports_punctuation_spacing_before_general_double_space_cleanup(self):
        source = read("src/forms/frmSpacingCleanup.bas")
        apply_block = source[source.index("' Perform replacements with counting"):]

        self.assertLess(apply_block.index("If doAfterPunct Then"), apply_block.index("If doDouble Then"))
        self.assertIn("cntBefore = cntBefore + ReplaceSpacingLiteralMatches(targetRange, CStr(beforePunct), Right$(CStr(beforePunct), 1))", source)
        self.assertIn('replacedThisPass = ReplaceSpacingLiteralMatches(targetRange, CStr(pattern), Left$(CStr(pattern), 1) & " ")', source)
        self.assertNotIn("With targetRange.Find", apply_block[: apply_block.index("RemoveAllHighlighting ActiveDocument.Content")])

    def test_spacing_tool_collapses_body_blanks_but_never_deletes_table_rows(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        self.assertIn("CollectCollapsibleBlankParagraphStarts(targetRange, False", source)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(ActiveDocument, CLng(blankStarts(blankIndex)), cbpkRemovableBody)", source)
        self.assertNotIn("CollapseExtraBlankTableRowsInRange", source)
        self.assertNotIn("IsSpacingBlankTableRow", source)
        self.assertNotIn(".Rows(", source)
        self.assertNotIn(".Rows.Delete", source)
        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", source)
        self.assertNotIn("RemoveAllHighlighting targetRange", source)

    def test_spacing_blank_line_preview_uses_shared_minimal_markers(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        self.assertIn("Set previewBlankStarts = CollectCollapsibleBlankParagraphStarts(targetRange, False", source)
        self.assertIn("previewBlank = previewBlankStarts.Count", source)
        self.assertIn("ApplyPreviewMinimalMarker ActiveDocument.Range(previewBlankStarts(previewBlankIndex), previewBlankStarts(previewBlankIndex))", source)
        self.assertNotIn("HighlightExtraBlankLinesInRange", source)
        self.assertNotIn("ApplyPreviewShading p.Range", source)
        self.assertNotIn('If doBlank Then previewBlank = CountPreviewFindMatches(targetRange, "^p^p^p")', source)

    def test_hyperlink_cleaner_clears_preview_highlighting_globally_after_apply(self):
        source = read("src/forms/frmHyperlinkRemover.bas")
        apply_block = source[source.index('MarkCleanupStart "Hyperlink Cleaner"'):source.index("MarkCleanupEnd", source.index('MarkCleanupStart "Hyperlink Cleaner"'))]

        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", apply_block)
        self.assertLess(
            apply_block.index("RemoveAllHighlighting ActiveDocument.Content"),
            apply_block.index("undoRec.EndCustomRecord"),
        )

    def test_hyperlink_remove_mode_does_not_rebuild_range_after_delete(self):
        source = read("src/forms/frmHyperlinkRemover.bas")
        apply_block = source[source.index('MarkCleanupStart "Hyperlink Cleaner"'):source.index("RemoveAllHighlighting ActiveDocument.Content")]
        remove_block = apply_block[apply_block.index("If removeLinks Then"):apply_block.index("Else", apply_block.index("If removeLinks Then"))]

        self.assertIn("Set rg = hl.Range.Duplicate", remove_block)
        self.assertLess(remove_block.index("Set rg = hl.Range.Duplicate"), remove_block.index("hl.Delete"))
        self.assertNotIn("ActiveDocument.Range(rs, rEnd)", remove_block)

    def test_hyperlink_visible_style_cleanup_preserves_font_size_and_face(self):
        source = read("src/forms/frmHyperlinkRemover.bas")
        helper = source[source.index("Private Sub ClearHyperlinkVisibleStyle"):source.index("End Sub", source.index("Private Sub ClearHyperlinkVisibleStyle"))]

        self.assertIn("target.Font.Underline = wdUnderlineNone", helper)
        self.assertIn("target.Font.ColorIndex = wdAuto", helper)
        self.assertNotIn("target.Style", helper)
        self.assertNotIn("wdStyleDefaultParagraphFont", helper)
        self.assertNotIn("Font.Size", helper)
        self.assertNotIn("Font.Name", helper)

    def test_list_cleanup_replaces_visible_paragraph_text_without_table_markers(self):
        source = read("src/forms/frmListCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("listText = Trim$(ParagraphBodyText(p))", source)
        self.assertIn("Set p = ReplaceAndRefreshListParagraph(p, Mid$(listText, 3))", source)
        self.assertIn("Set p = ReplaceAndRefreshListParagraph(p, LTrim$(Mid$(listText, pos + 1)))", source)
        self.assertIn("Private Function ReplaceAndRefreshListParagraph(ByVal p As Paragraph, ByVal replacementText As String) As Paragraph", source)
        self.assertIn("paragraphStart = p.Range.Start", source)
        self.assertIn("ReplaceParagraphBodyText p, replacementText", source)
        self.assertIn("Set ReplaceAndRefreshListParagraph = ActiveDocument.Range(paragraphStart, paragraphStart).Paragraphs(1)", source)
        self.assertIn('Mid$(fullText, bodyLength, 1) = vbCr Or Mid$(fullText, bodyLength, 1) = Chr$(7)', helpers)
        self.assertNotIn("p.Range.Text = Mid$(txt, 3) & vbCr", source)
        self.assertNotIn("p.Range.Text = LTrim$(Mid$(txt, pos + 1)) & vbCr", source)

    def test_list_cleanup_removes_manual_number_text_before_applying_word_numbering(self):
        source = read("src/forms/frmListCleanup.bas")
        numbering_block = source[source.index("Private Function NormalizeNumbering"):source.index("End Function", source.index("Private Function NormalizeNumbering"))]

        self.assertLess(
            numbering_block.index("Set p = ReplaceAndRefreshListParagraph(p, LTrim$(Mid$(listText, pos + 1)))"),
            numbering_block.index("ApplyListNumberStyle p, inNumberRun"),
        )

    def test_list_cleanup_restarts_manual_numbered_runs_at_first_item(self):
        source = read("src/forms/frmListCleanup.bas")
        numbering_block = source[source.index("Private Function NormalizeNumbering"):source.index("End Function", source.index("Private Function NormalizeNumbering"))]
        helper_block = source[source.index("Private Sub ApplyListNumberStyle"):source.index("End Sub", source.index("Private Sub ApplyListNumberStyle"))]

        self.assertIn("Dim inNumberRun As Boolean", numbering_block)
        self.assertIn("ApplyListNumberStyle p, inNumberRun", numbering_block)
        self.assertIn("inNumberRun = True", numbering_block)
        self.assertIn("If Len(listText) > 0 Then inNumberRun = False", numbering_block)
        self.assertIn("ApplyListTemplateWithLevel", helper_block)
        self.assertIn("ContinuePreviousList:=continuePreviousList", helper_block)
        self.assertNotIn("ApplyNumberDefault", numbering_block)

    def test_list_cleanup_tightens_list_paragraph_spacing_after_word_list_formatting(self):
        source = read("src/forms/frmListCleanup.bas")

        self.assertIn("TightenListParagraphSpacing p", source)
        self.assertIn("Private Sub TightenListParagraphSpacing(ByVal p As Paragraph)", source)
        self.assertIn(".SpaceBefore = 0", source)
        self.assertIn(".SpaceAfter = 0", source)
        self.assertIn(".LineSpacingRule = wdLineSpaceSingle", source)

    def test_list_cleanup_does_not_toggle_converted_bullets_back_off(self):
        source = read("src/forms/frmListCleanup.bas")
        run_block = source[source.index("Dim cntHyphen As Long"):source.index("RemoveAllHighlighting targetRange")]
        convert_block = source[source.index("Private Function ConvertHyphenListsToBullets"):source.index("End Function", source.index("Private Function ConvertHyphenListsToBullets"))]
        normalize_block = source[source.index("Private Function NormalizeBullets"):source.index("End Function", source.index("Private Function NormalizeBullets"))]
        helper_block = source[source.index("Private Sub ApplyListBulletStyle"):source.index("End Sub", source.index("Private Sub ApplyListBulletStyle"))]

        self.assertLess(
            run_block.index("If doBullets   Then cntBullets"),
            run_block.index("If doHyphen    Then cntHyphen"),
        )
        self.assertIn("ApplyListBulletStyle p", convert_block)
        self.assertIn("Private Sub ApplyListBulletStyle(ByVal p As Paragraph)", source)
        self.assertNotIn("ApplyListBulletStyle p", normalize_block)
        self.assertIn("If p.Range.ListFormat.ListType <> wdListBullet Then", helper_block)
        self.assertIn("p.Range.ListFormat.ApplyBulletDefault", helper_block)
        self.assertNotIn("p.Range.Style = wdStyleListBullet", helper_block)

    def test_list_cleanup_preserves_paragraph_shading_and_uses_hanging_list_indent(self):
        source = read("src/forms/frmListCleanup.bas")
        convert_block = source[source.index("Private Function ConvertHyphenListsToBullets"):source.index("End Function", source.index("Private Function ConvertHyphenListsToBullets"))]
        numbering_block = source[source.index("Private Function NormalizeNumbering"):source.index("End Function", source.index("Private Function NormalizeNumbering"))]
        indent_helper = source[source.index("Private Sub ApplyStandardListIndent"):source.index("End Sub", source.index("Private Sub ApplyStandardListIndent"))]
        fix_indent_block = source[source.index("Private Function FixListIndentation"):source.index("End Function", source.index("Private Function FixListIndentation"))]

        self.assertIn("shadingColor = p.Range.Shading.BackgroundPatternColor", convert_block)
        self.assertIn("RestoreListParagraphShading p, shadingColor", convert_block)
        self.assertIn("shadingColor = p.Range.Shading.BackgroundPatternColor", numbering_block)
        self.assertIn("RestoreListParagraphShading p, shadingColor", numbering_block)
        self.assertIn("Private Sub RestoreListParagraphShading(ByVal p As Paragraph, ByVal shadingColor As Long)", source)
        self.assertIn("Private Function ListIndentNeedsFix(ByVal p As Paragraph) As Boolean", source)
        self.assertIn("If ListIndentNeedsFix(p) Then", fix_indent_block)
        self.assertIn(".LeftIndent = InchesToPoints(0.5)", indent_helper)
        self.assertIn(".FirstLineIndent = InchesToPoints(-0.25)", indent_helper)
        self.assertNotIn("p.LeftIndent = InchesToPoints(0.25)", source)

    def test_list_cleanup_preview_and_apply_respect_selected_scope(self):
        source = read("src/forms/frmListCleanup.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("MarkCleanupStart \"List Normalizer\"")]
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("If ParagraphOverlapsRangeOutsideTable(p, targetRange) Then", preview_block)
        self.assertIn("If ParagraphOverlapsRangeOutsideTable(p, scopeRange) Then", source)
        self.assertIn("Public Function ParagraphOverlapsRangeOutsideTable", helpers)
        self.assertIn("If paragraphItem.Range.Information(wdWithInTable) Then Exit Function", helpers)
        self.assertNotIn("ListParagraphInScope", source)
        self.assertNotIn("If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then", source)

    def test_list_cleanup_preview_counts_match_apply_categories(self):
        source = read("src/forms/frmListCleanup.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("MarkCleanupStart \"List Normalizer\"")]

        self.assertIn("ResolveListParagraphActions p, txt, doBullets, doNumbering, doIndent, doHyphen", preview_block)
        self.assertIn("If actionHyphen Or actionNumbering Then HighlightManualListPrefix p, actionNumbering", preview_block)
        self.assertIn("Private Sub HighlightManualListPrefix", source)
        self.assertNotIn("p.Range.HighlightColorIndex = wdYellow", preview_block)
        self.assertIn("Private Sub ResolveListParagraphActions(", source)
        self.assertIn("If doIndent And ListIndentNeedsFix(p) Then actionIndent = True", source)
        self.assertNotIn('If doBullets And (Left$(txt, 2) = Chr(149) & " " Or Left$(txt, 2) = "* ")', preview_block)

    def test_paragraph_cleanup_uses_shared_revalidated_blank_classifier(self):
        source = read("src/forms/frmParagraphCleanup.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("MarkCleanupStart \"Paragraph Fixer\"")]
        apply_block = source[source.index("Dim cntRemove As Long"):source.index("RemoveAllHighlighting")]
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Set targetRange = ParagraphCleanupTargetRange(targetRange)", source)
        self.assertIn("Private Function ParagraphCleanupTargetRange(ByVal sourceRange As Range) As Range", source)
        self.assertIn("expanded.Start = expanded.Paragraphs(1).Range.Start", source)
        self.assertIn("expanded.End = expanded.Paragraphs(expanded.Paragraphs.Count).Range.End", source)
        self.assertIn("Public Function ParagraphOverlapsRange(ByVal paragraphItem As Paragraph, ByVal scopeRange As Range) As Boolean", helpers)
        self.assertIn("CollectCollapsibleBlankParagraphStarts(targetRange, True", preview_block)
        self.assertIn("ApplyPreviewMinimalMarker ActiveDocument.Range(previewBlankStarts(previewBlankIndex), previewBlankStarts(previewBlankIndex))", preview_block)
        self.assertIn("CollectCollapsibleBlankParagraphStarts(targetRange, True", apply_block)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(ActiveDocument, CLng(blankStarts(blankIndex)))", apply_block)
        self.assertNotIn('ReplaceParagraphFindMatches(targetRange, "^p^p^p", "^p^p")', apply_block)
        self.assertNotIn("p.Range.Delete", source)
        self.assertNotIn("p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End", source)

    def test_paragraph_cleanup_indent_count_uses_paragraph_format_object(self):
        source = read("src/forms/frmParagraphCleanup.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("MarkCleanupStart \"Paragraph Fixer\"")]
        apply_block = source[source.index("' Fix indent"):source.index("RemoveAllHighlighting")]

        self.assertIn("p.Format.LeftIndent", preview_block)
        self.assertIn("p.Format.LeftIndent", apply_block)
        self.assertNotIn("p.LeftIndent", source)

    def test_paragraph_cleanup_soft_return_option_converts_manual_line_breaks(self):
        source = read("src/forms/frmParagraphCleanup.bas")
        apply_block = source[source.index("Dim cntRemove As Long"):source.index("RemoveAllHighlighting")]

        self.assertIn('replacedThisPass = ReplaceParagraphFindMatches(targetRange, "^l", "^p")', apply_block)
        self.assertIn("Private Function ReplaceParagraphFindMatches(", source)
        self.assertIn(".Replacement.Text = replacementText", source)
        self.assertIn("Do While .Execute(Replace:=wdReplaceOne)", source)

    def test_soft_return_converter_excludes_required_final_paragraph_mark_from_counts(self):
        source = read("src/forms/frmSoftReturnConverter.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("ShowPreviewActionsSummary Me, \"Soft Return Converter\", previewRows")]
        apply_block = source[source.index('MarkCleanupStart "Soft Return Converter"'):source.index("undoRec.EndCustomRecord")]
        self.assertIn("Private Function CountSoftReturnConvertibleMarks(ByVal searchRange As Range, ByVal softToPara As Boolean) As Long", source)
        helper = source[source.index("Private Function CountSoftReturnConvertibleMarks"):source.index("End Function", source.index("Private Function CountSoftReturnConvertibleMarks"))]

        self.assertIn("pc = CountSoftReturnConvertibleMarks(targetRange, softToPara)", preview_block)
        self.assertIn("cnt = ReplaceSoftReturnMarks(targetRange, softToPara)", apply_block)
        self.assertIn("Private Function ReplaceSoftReturnMarks(ByVal searchRange As Range, ByVal softToPara As Boolean) As Long", source)
        replace_helper = source[source.index("Private Function ReplaceSoftReturnMarks"):source.index("End Function", source.index("Private Function ReplaceSoftReturnMarks"))]
        self.assertIn("If softToPara Then", helper)
        self.assertIn("CountSoftReturnCharacters(searchRange.Text, Chr$(11))", helper)
        self.assertIn("CountSoftReturnCharacters(searchRange.Text, Chr$(13))", helper)
        self.assertIn("If Right$(searchRange.Text, 1) = Chr$(13) Then", helper)
        self.assertIn("CountSoftReturnConvertibleMarks = CountSoftReturnConvertibleMarks - 1", helper)
        self.assertIn("If matchRange.End >= ActiveDocument.Content.End Then Exit Do", replace_helper)
        self.assertIn("matchRange.Text = Chr$(11)", replace_helper)
        self.assertIn("matchRange.Text = Chr$(13)", replace_helper)
        self.assertIn("If (Not softToPara) And matchRange.End >= searchEnd Then Exit Do", replace_helper)
        self.assertIn("Set matchRange = ActiveDocument.Range(nextStart, searchEnd)", replace_helper)
        self.assertNotIn(".Execute Replace:=wdReplaceAll", apply_block)
        self.assertIn("Private Function CountSoftReturnCharacters(ByVal textValue As String, ByVal markChar As String) As Long", source)
        self.assertNotIn("Len(pt) - Len(Replace(pt, markChar, \"\"))", source)
        self.assertNotIn("Len(txt) - Len(Replace(txt, markChar, \"\"))", source)

    def test_soft_return_preview_marks_nearest_visible_characters(self):
        source = read("src/forms/frmSoftReturnConverter.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("ShowPreviewActionsSummary Me, \"Soft Return Converter\", previewRows")]

        self.assertIn("pc = CountSoftReturnConvertibleMarks(targetRange, softToPara)", preview_block)
        self.assertIn("Call HighlightSoftReturnConvertibleMarks(targetRange, softToPara)", preview_block)
        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", preview_block)
        self.assertLess(preview_block.index("RemoveAllHighlighting ActiveDocument.Content"), preview_block.index("pc = CountSoftReturnConvertibleMarks(targetRange, softToPara)"))
        self.assertNotIn(".Replacement.Highlight = True", preview_block)
        self.assertNotIn("Private Function SoftReturnPreviewMarksAreVisible", source)
        self.assertIn("Private Function HighlightSoftReturnConvertibleMarks", source)
        self.assertIn("ApplyPreviewMinimalMarker matchRange", source)

    def test_soft_return_preview_marks_active_rows_highlighted_and_inactive_rows_gray(self):
        source = read("src/forms/frmSoftReturnConverter.bas")
        preview_block = source[source.index("If previewOnly Then"):source.index("ShowPreviewActionsSummary Me, \"Soft Return Converter\", previewRows")]

        self.assertIn('AddPreviewSummaryRow previewRows, "Soft returns", IIf(softToPara, pc, 0), False, (Not softToPara)', preview_block)
        self.assertIn('AddPreviewSummaryRow previewRows, "Paragraph marks", IIf(softToPara, 0, pc), False, softToPara', preview_block)

    def test_soft_return_converter_apply_clears_preview_highlighting_globally(self):
        source = read("src/forms/frmSoftReturnConverter.bas")
        apply_block = source[source.index('MarkCleanupStart "Soft Return Converter"'):source.index("MarkCleanupEnd", source.index('MarkCleanupStart "Soft Return Converter"'))]

        self.assertIn("RemoveAllHighlighting ActiveDocument.Content", apply_block)
        self.assertLess(
            apply_block.index("RemoveAllHighlighting ActiveDocument.Content"),
            apply_block.index("undoRec.EndCustomRecord"),
        )

    def test_duplicate_fuzzy_jaccard_does_not_use_reserved_shared_identifier(self):
        source = read("src/forms/frmDuplicateDetector.bas")
        jaccard = source[source.index("Private Function JaccardSimilarity"):source.index("End Function", source.index("Private Function JaccardSimilarity"))]

        self.assertIn("Dim sharedCount As Long", jaccard)
        self.assertIn("sharedCount = sharedCount + 1", jaccard)
        self.assertIn("JaccardSimilarity = sharedCount / (a.Count + b.Count - sharedCount)", jaccard)
        self.assertNotIn("Dim shared As Long", jaccard)

    def test_duplicate_detector_does_not_treat_image_only_paragraphs_as_duplicates(self):
        source = read("src/forms/frmDuplicateDetector.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        self.assertIn("candidateText = DuplicateCandidateVisibleText(p.Range.Text)", source)
        self.assertIn("If Len(candidateText) > 0 Then", source)
        self.assertIn("blankKind = ClassifyCleanupBlankParagraph(p)", source)
        self.assertIn("Case cbpkRemovableBody, cbpkRemovableCellExtra", source)
        self.assertIn("If targetRange.InlineShapes.Count > 0 Then GoTo MeaningfulRange", helpers)
        self.assertIn("paragraphText = Replace(paragraphText, Chr$(1), \"\")", source)
        self.assertIn("If a.Count = 0 And b.Count = 0 Then JaccardSimilarity = 0", source)

    def test_duplicate_remover_is_removal_only_and_safely_includes_empty_paragraphs(self):
        source = read("src/forms/frmDuplicateDetector.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn('chkIncludeEmptyParagraphs.Caption = "Include empty paragraphs"', source)
        self.assertIn("Dim doRemove As Boolean: doRemove = True", source)
        self.assertNotIn("optHighlightOnly.GroupName", source)
        self.assertNotIn("optRemoveDupes.GroupName", source)
        self.assertIn("blankKind = ClassifyCleanupBlankParagraph(paragraphItem)", source)
        self.assertIn("RemoveDuplicateParagraphAtStart = RemoveCleanupBlankParagraphAtStart(ActiveDocument, paragraphStart, blankKind)", source)
        self.assertIn("If expectedKind <> cbpkNotBlank And blankKind <> expectedKind Then Exit Function", helpers)
        self.assertIn("If paragraphItem.Range.Start <= cellRange.Start Then Exit Function", helpers)
        self.assertIn("If precedingMark.Text <> vbCr Then Exit Function", helpers)
        self.assertIn("If Not precedingMark.Information(wdWithInTable) Then Exit Function", helpers)
        self.assertIn("ApplyPreviewMinimalMarker paragraphItem.Range, False, wdYellow", source)
        self.assertIn("If RemoveDuplicateParagraphAtStart(ds(di)) Then removed = removed + 1", source)
        self.assertIn("Private Function RemoveDuplicateParagraphAtStart(ByVal paragraphStart As Long) As Boolean", source)
        self.assertIn("precedingMark.Delete", helpers)
        self.assertIn('AddPreviewSummaryRow previewRows, "Removable empty paragraphs", duplicateEmptyCount', source)
        self.assertIn('AddPreviewSummaryRow previewRows, "Protected empty cells", protectedEmptyCells', source)
        self.assertIn('AddPreviewSummaryRow previewRows, "Protected row markers", protectedRowMarkers', source)
        self.assertIn('Case "frmDuplicateDetector": GuidedToolName = "Duplicate Paragraph Remover"', helpers)
        self.assertIn('Case "frmDuplicateDetector": FriendlyFormCaption = "Duplicate Paragraph Remover"', installer)

    def test_duplicate_remover_fuzzy_thresholds_match_their_labels(self):
        source = read("src/forms/frmDuplicateDetector.bas")
        self.assertIn("Dim threshold As Double: threshold = 0.7", source)
        self.assertIn("If optFuzzyLoose.Value Then threshold = 0.5", source)
        self.assertIn("If optFuzzyStrict.Value Then threshold = 0.9", source)

    def test_formatting_stripper_uses_supported_font_reset_api(self):
        source = read("src/forms/frmFormattingStripper.bas")
        self.assertEqual(3, source.count("p.Range.Font.Reset"))
        self.assertNotIn("ClearCharacterDirectFormatting", source)

    def test_footnote_inline_text_is_inserted_after_reference_deletion(self):
        source = read("src/forms/frmFootnoteRemover.bas")
        apply_start = source.index('MarkCleanupStart "Footnote / Endnote Remover"')
        foot_start = source.index("If doFoot Then", apply_start)
        end_start = source.index("If doEnd Then", foot_start)
        foot_block = source[foot_start:end_start]
        end_block = source[end_start:source.index("undoRec.EndCustomRecord", end_start)]
        self.assertLess(foot_block.index("fn.Delete"), foot_block.index("ActiveDocument.Range(footInsertAt, footInsertAt).InsertAfter"))
        self.assertLess(end_block.index("en.Delete"), end_block.index("ActiveDocument.Range(endInsertAt, endInsertAt).InsertAfter"))
        self.assertIn("Private Function CleanInlineNoteText", source)

    def test_metadata_and_final_review_split_responsibilities(self):
        metadata = read("src/modules/modMetaDataSuite.bas")
        final_review = read("src/forms/frmFinalReview.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertNotIn("chkComments", metadata)
        self.assertNotIn("chkRevisions", metadata)
        self.assertIn("chkComments", final_review)
        self.assertIn("chkRevisions", final_review)
        self.assertNotIn('If toolForm.Name = "frmMetadataScrubber" Then', helpers)
        self.assertNotIn('Case "chkProperties", "chkPersonalInfo"', helpers)

    def test_document_trim_marks_apply_exit_before_unloading_on_no_op(self):
        source = read("src/forms/frmDocumentTrim.bas")

        self.assertIn("If toRemove <= 0 Then", source)
        no_op_block = source[source.index("If toRemove <= 0 Then"):source.index("End If", source.index("If toRemove <= 0 Then"))]
        self.assertIn("MarkCleanupToolApplied", no_op_block)
        self.assertIn('MsgBox "No trailing empty paragraphs found.", vbInformation', no_op_block)
        self.assertIn("Unload Me", no_op_block)
