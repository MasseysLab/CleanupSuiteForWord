from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class PreviewInactiveRowTests(unittest.TestCase):
    def test_paragraph_preview_find_moves_forward_without_replacing_matches(self):
        source = read("src/forms/frmParagraphCleanup.bas")
        helper_start = source.index("Private Function HighlightParagraphFindMatches")
        helper_end = source.index("End Function", helper_start)
        helper = source[helper_start:helper_end]

        self.assertIn("ApplyPreviewHighlight matchRange, wdYellow", helper)
        self.assertIn("matchRange.Collapse wdCollapseEnd", helper)
        self.assertIn("matchRange.End = searchEnd", helper)
        self.assertNotIn(".Replacement.Highlight = True", helper)
        self.assertNotIn("Replace:=wdReplaceOne", helper)
        self.assertIn('cntCollapse = HighlightPreviewFindBoundaryMatches(targetRange, "^l")', source)

    def test_mode_driven_previews_gray_only_unselected_rows(self):
        expected_calls = {
            "src/forms/frmBreakNormalizer.bas": [
                'AddPreviewSummaryRow previewRows, "Consecutive section breaks", previewSectionBreaks, False, (Not doCollSect)',
                'AddPreviewSummaryRow previewRows, "Consecutive page breaks", previewPageBreaks, False, (Not doCollPage)',
                'AddPreviewSummaryRow previewRows, "Section break conversions", previewConversions, False, (Not doConvert)',
            ],
            "src/forms/frmFinalReview.bas": [
                'AddPreviewSummaryRow previewRows, "Comments", previewComments, False, (Not doComments)',
                'AddPreviewSummaryRow previewRows, "Tracked changes", previewRevisions, False, (Not doRevisions)',
            ],
            "src/forms/frmFontNormalizer.bas": [
                'AddPreviewSummaryRow previewRows, "Font face", previewFace, False, (Not doFace)',
                'AddPreviewSummaryRow previewRows, "Font size", previewSize, False, (Not doSize)',
                'AddPreviewSummaryRow previewRows, "Bold", previewBold, False, (Not doBold)',
                'AddPreviewSummaryRow previewRows, "Italic", previewItalic, False, (Not doItalic)',
                'AddPreviewSummaryRow previewRows, "Font color", previewColor, False, (Not doColor)',
            ],
            "src/forms/frmFootnoteRemover.bas": [
                'AddPreviewSummaryRow previewRows, "Footnotes", pf, False, (Not doFoot)',
                'AddPreviewSummaryRow previewRows, "Endnotes", pe, False, (Not doEnd)',
                'AddPreviewSummaryRow previewRows, "Note text kept inline", keptInline, False, (Not keepText)',
            ],
            "src/forms/frmFormattingStripper.bas": [
                'AddPreviewSummaryRow previewRows, "Character formatting", previewChar, False, (Not doResetChar)',
                'AddPreviewSummaryRow previewRows, "Paragraph formatting", previewPara, False, (Not doResetPara)',
            ],
            "src/forms/frmHeaderFooterStandardizer.bas": [
                'AddPreviewSummaryRow previewRows, "Header areas", headerAreas, False, (Not doHeaders)',
                'AddPreviewSummaryRow previewRows, "Footer areas", footerAreas, False, (Not doFooters)',
                'AddPreviewSummaryRow previewRows, "Cleared areas", previewCleared, False, (Not doClear)',
                'AddPreviewSummaryRow previewRows, "Font reset areas", previewFont, False, (Not standardizeFont)',
                'AddPreviewSummaryRow previewRows, "Spacing reset areas", previewSpacing, False, (Not standardizeSpacing)',
                'AddPreviewSummaryRow previewRows, "Alignment areas", previewAlignment, False, (Not standardizeAlignment)',
                'AddPreviewSummaryRow previewRows, "Unlinked sections", previewUnlinked, False, (Not breakLinks)',
            ],
            "src/forms/frmListCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Hyphen list items", previewHyphen, False, (Not doHyphen)',
                'AddPreviewSummaryRow previewRows, "Bullet list items", previewBullets, False, (Not doBullets)',
                'AddPreviewSummaryRow previewRows, "Numbered list items", previewNumbering, False, (Not doNumbering)',
                'AddPreviewSummaryRow previewRows, "List indents", previewIndent, False, (Not doIndent)',
            ],
            "src/forms/frmMetadataScrubber.bas": [
                'AddPreviewSummaryRow previewRows, "Document properties", previewProperties, True, (Not doProps)',
                'AddPreviewSummaryRow previewRows, "Personal information", previewPersonal, True, (Not doPersonal)',
            ],
            "src/forms/frmObjectRemover.bas": [
                'AddPreviewSummaryRow previewRows, "Pictures", previewPictures, False, (Not doPic)',
                'AddPreviewSummaryRow previewRows, "Text boxes", previewTextBoxes, False, (Not doTxt)',
                'AddPreviewSummaryRow previewRows, "Frames", previewFrames, False, (Not doFra)',
                'AddPreviewSummaryRow previewRows, "Horizontal lines", previewHorizontalLines, False, (Not doHL)',
                'AddPreviewSummaryRow previewRows, "Form controls", previewFormControls, False, (Not doHtml)',
                'AddPreviewSummaryRow previewRows, "Hidden text", previewHiddenText, False, (Not doHid)',
                'AddPreviewSummaryRow previewRows, "Tables", previewTables, False, (Not doTab)',
            ],
            "src/forms/frmParagraphCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Empty paragraphs", cntRemove, False, (Not doRemove)',
                'AddPreviewSummaryRow previewRows, "Soft returns", cntCollapse, False, (Not doCollapse)',
                'AddPreviewSummaryRow previewRows, "Paragraph spacing", cntNormalize, False, (Not doNormalize)',
                'AddPreviewSummaryRow previewRows, "Indents", cntIndent, False, (Not doIndent)',
            ],
            "src/forms/frmSpacingCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Double spaces", previewDouble, False, (Not doDouble)',
                'AddPreviewSummaryRow previewRows, "Leading/trailing spaces", previewTrim, False, (Not doTrim)',
                'AddPreviewSummaryRow previewRows, "Spaces before punctuation", previewBefore, False, (Not doBeforePunct)',
                'AddPreviewSummaryRow previewRows, "Spaces after punctuation", previewAfter, False, (Not doAfterPunct)',
                'AddPreviewSummaryRow previewRows, "Extra blank lines", previewBlank, False, (Not doBlank)',
            ],
            "src/forms/frmStyleCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Unused custom styles", previewUnused, True, (Not doUnused)',
                'AddPreviewSummaryRow previewRows, "Variant style remaps", previewRemaps, False, (Not doRemap)',
            ],
            "src/forms/frmTableCleaner.bas": [
                'AddPreviewSummaryRow previewRows, "Empty rows", previewEmptyRows, False, (Not doEmptyRows)',
                'AddPreviewSummaryRow previewRows, "Empty columns", previewCols, False, (Not doEmptyCols)',
                'AddPreviewSummaryRow previewRows, "Cell padding", previewPadding, False, (Not doPadding)',
                'AddPreviewSummaryRow previewRows, "Direct cell formatting", previewDirectFmt, False, (Not doDirectFmt)',
                'AddPreviewSummaryRow previewRows, "Border normalization", previewNormalizeBorders, False, (Not doBorders)',
                'AddPreviewSummaryRow previewRows, "Border removal", previewRemoveBorders, False, (Not doRemoveBorders)',
                'AddPreviewSummaryRow previewRows, "Tables to text", previewConvertText, False, (Not doConvertText)',
            ],
        }

        for path, calls in expected_calls.items():
            source = read(path)
            for call in calls:
                with self.subTest(path=path, call=call):
                    self.assertIn(call, source)

    def test_category_previews_use_explicit_selection_flags(self):
        expected_calls = {
            "src/forms/frmPunctuationCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Curly double quotes", previewCurlyDouble, False, (Not includeCurlyDouble)',
                'AddPreviewSummaryRow previewRows, "Curly single quotes", previewCurlySingle, False, (Not includeCurlySingle)',
                'AddPreviewSummaryRow previewRows, "Em dashes", previewEmDash, False, (Not includeEmDash)',
                'AddPreviewSummaryRow previewRows, "En dashes", previewEnDash, False, (Not includeEnDash)',
                'AddPreviewSummaryRow previewRows, "Ellipses", previewEllipses, False, (Not includeEllipses)',
            ],
            "src/forms/frmUnicodeCleanup.bas": [
                'AddPreviewSummaryRow previewRows, "Non-breaking spaces", previewNBSP, False, (Not includeNBSP)',
                'AddPreviewSummaryRow previewRows, "Zero-width spaces", previewZWSP, False, (Not includeZWSP)',
                'AddPreviewSummaryRow previewRows, "Zero-width non-joiners", previewZWNJ, False, (Not includeZWNJ)',
                'AddPreviewSummaryRow previewRows, "Zero-width joiners", previewZWJ, False, (Not includeZWJ)',
                'AddPreviewSummaryRow previewRows, "Byte order marks", previewBOM, False, (Not includeBOM)',
                'AddPreviewSummaryRow previewRows, "Soft hyphens", previewSoftHyphen, False, (Not includeSoftHyphen)',
                'AddPreviewSummaryRow previewRows, "Non-breaking hyphens", previewNBHyphen, False, (Not includeNBHyphen)',
            ],
        }

        for path, calls in expected_calls.items():
            source = read(path)
            for call in calls:
                with self.subTest(path=path, call=call):
                    self.assertIn(call, source)

    def test_large_preview_scans_skip_inactive_options(self):
        object_source = read("src/forms/frmObjectRemover.bas")
        metadata_source = read("src/forms/frmMetadataScrubber.bas")
        style_source = read("src/forms/frmStyleCleanup.bas")

        self.assertIn("If doPic Then previewPictures = ProcessPictures(False)", object_source)
        self.assertNotIn("IIf(doPic, ProcessPictures(False), 0)", object_source)
        self.assertIn("If doProps Then previewProperties = CountPopulatedMetadataValues(propertySignals)", metadata_source)
        self.assertNotIn("IIf(doProps, CountPopulatedMetadataValues(propertySignals), 0)", metadata_source)
        self.assertIn("If doRemap Then previewRemaps = HighlightStyleVariantsForRemap()", style_source)
        self.assertNotIn("IIf(doRemap, HighlightStyleVariantsForRemap(), 0)", style_source)

    def test_break_preview_marks_nearest_visible_characters_without_replacing(self):
        source = read("src/forms/frmBreakNormalizer.bas")
        preview_start = source.index("If previewOnly Then")
        preview_end = source.index('MarkCleanupStart "Break Normalizer"', preview_start)
        preview_block = source[preview_start:preview_end]

        self.assertIn('HighlightPreviewFindBoundaryMatches(targetRange, "^b^b")', preview_block)
        self.assertIn('HighlightPreviewFindBoundaryMatches(targetRange, "^m^m")', preview_block)
        self.assertIn("ApplyPreviewMinimalMarker previewSec.Range, True", preview_block)
        self.assertNotIn(".Replacement.Highlight = True", preview_block)
        self.assertNotIn("Replace:=wdReplaceAll", preview_block)


if __name__ == "__main__":
    unittest.main()
