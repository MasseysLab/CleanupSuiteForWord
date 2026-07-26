from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class StructuralSafetyContractTests(unittest.TestCase):
    def test_shared_blank_classifier_is_fail_closed_and_preserves_word_markers(self):
        source = read("src/modules/modCleanupHelpers.bas")

        for name in (
            "cbpkRemovableBody",
            "cbpkRemovableCellExtra",
            "cbpkProtectedEmptyCell",
            "cbpkProtectedRowMarker",
            "cbpkProtectedFinalDocument",
            "cbpkUnsupported",
        ):
            self.assertIn(name, source)

        self.assertIn("If targetRange.InlineShapes.Count > 0 Then GoTo MeaningfulRange", source)
        self.assertIn("If targetRange.Fields.Count > 0 Then GoTo MeaningfulRange", source)
        self.assertIn("If targetRange.FormFields.Count > 0 Then GoTo MeaningfulRange", source)
        self.assertIn("If targetRange.ContentControls.Count > 0 Then GoTo MeaningfulRange", source)
        self.assertIn("If CleanupRangeHasUserBookmark(targetRange) Then GoTo MeaningfulRange", source)
        self.assertIn("If targetRange.Revisions.Count > 0 Then GoTo MeaningfulRange", source)
        self.assertIn("If CleanupRangeHasAnchoredShape(targetRange) Then GoTo MeaningfulRange", source)
        self.assertNotIn('rawText = Replace(rawText, Chr$(160), "")', source)
        self.assertIn("If CleanupCellRangeContainsNestedTable(cellRange) Then GoTo UnsupportedParagraph", source)
        self.assertIn("If CleanupCellRangeContainsNestedTable(tableCell.Range) Then CleanupCellHasMeaningfulContent = True", source)
        self.assertIn("rawText = cellRange.Text", source)
        self.assertIn("If cellMarkerCount > 1 Then", source)
        self.assertNotIn("tableCell.NestingLevel", source)
        self.assertIn("ProtectRange:\n    CleanupRangeHasMeaningfulContent = True", source)

    def test_body_blank_tools_share_collection_and_apply_revalidation(self):
        spacing = read("src/forms/frmSpacingCleanup.bas")
        paragraph = read("src/forms/frmParagraphCleanup.bas")
        trim = read("src/forms/frmDocumentTrim.bas")

        self.assertIn("CollectCollapsibleBlankParagraphStarts(targetRange, False", spacing)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(ActiveDocument", spacing)
        self.assertNotIn(".Rows.Delete", spacing)
        self.assertIn("CollectCollapsibleBlankParagraphStarts(targetRange, True", paragraph)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(ActiveDocument", paragraph)
        self.assertIn("blankKind = ClassifyCleanupBlankParagraph(paragraphItem)", trim)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(ActiveDocument", trim)

    def test_table_structure_actions_are_opt_in_previewed_confirmed_and_revalidated(self):
        source = read("src/forms/frmTableCleaner.bas")

        self.assertIn("chkRemoveEmptyRows.Value = False", source)
        self.assertIn("chkRemoveEmptyCols.Value = False", source)
        self.assertIn("chkConvertToText.Value = False", source)
        self.assertIn("If (doEmptyRows Or doEmptyCols Or doConvertText) And Not mStructuralPreviewCompleted Then", source)
        self.assertIn("vbExclamation + vbYesNo + vbDefaultButton2", source)
        self.assertIn("CollectEligibleEmptyRowIndexes", source)
        self.assertIn("CollectEligibleEmptyColumnIndexes", source)
        self.assertIn("TableRowStillEligible", source)
        self.assertIn("TableColumnStillEligible", source)
        self.assertIn("For rowCandidateIndex = applyRowIndexes.Count To 1 Step -1", source)
        self.assertIn("For columnCandidateIndex = applyColumnIndexes.Count To 1 Step -1", source)
        self.assertIn("convertThisTable = (doConvertText And tbl.Columns.Count = 1 And CleanupTableStructureIsSupported(tbl))", source)
        self.assertIn("If convertThisTable Then", source)
        self.assertIn("undoRec.StartCustomRecord \"Cleanup Suite - Table Cleaner\"", source)
        apply_start = source.index('undoRec.StartCustomRecord "Cleanup Suite - Table Cleaner"')
        apply_end = source.index("Unload Me", apply_start)
        apply_block = source[apply_start:apply_end]
        self.assertLess(apply_block.index("MarkCleanupEnd"), apply_block.index("undoRec.EndCustomRecord"))
        self.assertLess(apply_block.index("MarkCleanupToolApplied"), apply_block.index("undoRec.EndCustomRecord"))

    def test_table_classifier_protects_ambiguous_and_total_structure_deletion(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        table = read("src/forms/frmTableCleaner.bas")

        self.assertNotIn("If Not sourceTable.Uniform Then Exit Function", helpers)
        self.assertIn("For rowIndex = 1 To sourceTable.Rows.Count", helpers)
        self.assertIn("For columnIndex = 1 To sourceTable.Columns.Count", helpers)
        self.assertIn("If CleanupCellRangeContainsNestedTable(sourceTable.Cell(rowIndex, columnIndex).Range) Then Exit Function", helpers)
        self.assertIn("If sourceTable.Rows(rowIndex).HeadingFormat <> 0 Then Exit Function", table)
        self.assertIn("If indexes.Count >= sourceTable.Rows.Count Then", table)
        self.assertIn("If indexes.Count >= sourceTable.Columns.Count Then", table)
        self.assertIn('AddPreviewSummaryRow previewRows, "Protected/skipped rows"', table)
        self.assertIn('AddPreviewSummaryRow previewRows, "Protected/skipped columns"', table)
        self.assertIn('AddPreviewSummaryRow previewRows, "Skipped multi-column tables"', table)
        self.assertIn('AddPreviewSummaryRow previewRows, "Protected/unsupported conversions"', table)
        self.assertIn('AddPreviewSummaryRow previewRows, "Partially scoped tables"', table)

    def test_break_collapse_structural_choices_default_off(self):
        source = read("src/forms/frmBreakNormalizer.bas")

        self.assertIn("chkCollapseSectionBreaks.Value = False", source)
        self.assertIn("chkCollapsePageBreaks.Value = False", source)
        self.assertNotIn("chkCollapseSectionBreaks.Value = True", source)
        self.assertNotIn("chkCollapsePageBreaks.Value = True", source)

    def test_word_smoke_hook_exercises_real_structural_guards(self):
        source = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn("Public Function RunStructuralSafetySmokeCheck() As String", source)
        self.assertIn("ClassifyCleanupBlankParagraph(testDocument.Paragraphs(2)) = cbpkRemovableBody", source)
        self.assertIn("CollectCollapsibleBlankParagraphStarts(testDocument.Content, False", source)
        self.assertIn("RemoveCleanupBlankParagraphAtStart(testDocument, removableStart, cbpkRemovableBody)", source)
        self.assertIn("Not CleanupTableRowHasMeaningfulContent(testTable.Rows(2))", source)
        self.assertIn('testTable.Cell(2, 1).Range.InsertBefore Chr$(160)', source)
        self.assertIn("CleanupCellHasMeaningfulContent(testTable.Cell(2, 2))", source)
        self.assertIn("Not CleanupTableStructureIsSupported(testTable)", source)
        self.assertIn("Not CleanupTableStructureIsSupported(mergedTable)", source)
        self.assertIn("Public Function RunStructuralFixtureSmokeCheck() As String", source)
        self.assertIn("CleanupTableStructureIsSupported(ActiveDocument.Tables(1))", source)
        self.assertIn("Not CleanupTableStructureIsSupported(ActiveDocument.Tables(5))", source)
        self.assertIn("Not CleanupTableStructureIsSupported(ActiveDocument.Tables(6))", source)


if __name__ == "__main__":
    unittest.main()
