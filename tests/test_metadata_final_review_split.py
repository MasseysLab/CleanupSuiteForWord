from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class MetadataFinalReviewSplitTests(unittest.TestCase):
    def test_metadatasuite_clear_sharing_properties_replaces_metadata_scrubber(self):
        source = read("src/modules/modMetaDataSuite.bas")

        for snippet in [
            'Public Sub ClearMetadataDashboardSharingProperties(ByVal frm As Object)',
            'previewText = BuildSharingPropertiesPreview(doc)',
            'Private Function BuildSharingPropertiesPreview(ByVal doc As Object) As String',
            'Private Sub ApplyClearSharingProperties(ByVal doc As Object)',
            'doc.RemoveDocumentInformation wdRDIDocumentProperties',
            'doc.RemoveDocumentInformation wdRDIRemovePersonalInformation',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, source)

    def test_final_review_preview_and_apply_use_explicit_selection_wording(self):
        source = read("src/forms/frmFinalReview.bas")

        for snippet in [
            'Set previewRows = NewPreviewSummaryRows()',
            'AddPreviewSummaryRow previewRows, "Comments", previewComments, True, (Not doComments)',
            'AddPreviewSummaryRow previewRows, "Tracked changes", previewRevisions, True, (Not doRevisions)',
            'ShowPreviewActionsSummary Me, "Final Review", previewRows',
            'ActiveDocument.DeleteAllComments',
            'ActiveDocument.Revisions.AcceptAll',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, source)

        self.assertNotIn("BuildFinalReviewPreviewSummary", source)
        self.assertNotIn("ShowCleanupReport", source)


if __name__ == "__main__":
    unittest.main()
