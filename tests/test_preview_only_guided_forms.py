from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


GUIDED_FORM_FILES = [
    "src/forms/frmPunctuationCleanup.bas",
    "src/forms/frmUnicodeCleanup.bas",
    "src/forms/frmSpacingCleanup.bas",
    "src/forms/frmCapitalizationCleanup.bas",
    "src/forms/frmListCleanup.bas",
    "src/forms/frmParagraphCleanup.bas",
    "src/forms/frmDuplicateDetector.bas",
    "src/forms/frmFontNormalizer.bas",
    "src/forms/frmTableCleaner.bas",
    "src/forms/frmBreakNormalizer.bas",
    "src/forms/frmDocumentTrim.bas",
    "src/forms/frmFormattingStripper.bas",
    "src/forms/frmHyperlinkRemover.bas",
    "src/forms/frmSoftReturnConverter.bas",
    "src/forms/frmFinalReview.bas",
    "src/forms/frmStyleCleanup.bas",
    "src/forms/frmFootnoteRemover.bas",
    "src/forms/frmHeaderFooterStandardizer.bas",
    "src/forms/frmObjectRemover.bas",
]


class PreviewOnlyGuidedFormTests(unittest.TestCase):
    def test_preview_actions_is_the_only_visible_apply_surface(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        preview_actions = read("src/forms/frmPreviewActions.bas")

        self.assertIn('cmdApply.Caption = "Apply"', preview_actions)
        self.assertIn("cmdApply.Move M + (2 * (BTN_W + GAP)), ACTION_TOP, BTN_W, buttonHeight", preview_actions)
        self.assertNotIn('toolForm.Controls("cmdRun").Visible = True', helpers)
        self.assertNotIn('toolForm.Controls("cmdRun").Caption = "Apply"', helpers)
        self.assertNotIn('PositionControl designer, "cmdRun", MARGIN, y, halfW, BTN_H', installer)

    def test_guided_forms_keep_internal_run_after_preview_plumbing(self):
        for form_path in GUIDED_FORM_FILES:
            source = read(form_path)
            with self.subTest(form=form_path):
                self.assertIn("Public Sub RunAfterPreview()", source)
                self.assertIn("cmdRun_Click", source)
                self.assertIn("ShowPreviewActionsSummary Me", source)

    def test_guided_forms_use_structured_preview_summary_rows(self):
        for form_path in GUIDED_FORM_FILES:
            source = read(form_path)
            with self.subTest(form=form_path):
                self.assertIn("Set previewRows = NewPreviewSummaryRows()", source)
                self.assertIn("ShowPreviewActionsSummary Me", source)
                self.assertNotIn('"Preview complete.', source)

    def test_preview_summaries_keep_zero_count_planned_rows(self):
        paragraph = read("src/forms/frmParagraphCleanup.bas")

        self.assertIn('AddPreviewSummaryRow previewRows, "Empty paragraphs", cntRemove', paragraph)
        self.assertIn('AddPreviewSummaryRow previewRows, "Soft returns", cntCollapse', paragraph)
        self.assertIn('AddPreviewSummaryRow previewRows, "Paragraph spacing", cntNormalize', paragraph)
        self.assertIn('AddPreviewSummaryRow previewRows, "Indents", cntIndent', paragraph)
        self.assertNotIn("If cntRemove > 0 Then AddPreviewSummaryRow", paragraph)
        self.assertNotIn("If cntCollapse > 0 Then AddPreviewSummaryRow", paragraph)
        self.assertNotIn("If cntNormalize > 0 Then AddPreviewSummaryRow", paragraph)
        self.assertNotIn("If cntIndent > 0 Then AddPreviewSummaryRow", paragraph)

    def test_metadatasuite_is_not_a_generic_guided_form(self):
        metadata_form = read("src/forms/frmMetaDataSuite.bas")
        self.assertIn("InitializeMetadataDashboard Me", metadata_form)
        self.assertNotIn("LayoutCleanupToolForm Me", metadata_form)

    def test_scoped_layout_contract_is_encoded_once(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")
        self.assertIn("Private Sub LayoutGuidedPreviewAndScopeRow", helpers)
        self.assertIn('"Selected text only"', helpers)
        self.assertIn('"Entire document"', helpers)
        self.assertIn('toolForm.Controls("cmdPreview").Move previewLeft, y, previewW, 42', helpers)
        self.assertIn('toolForm.Controls("cmdPreview").Move M, y, contentW, 30', helpers)
        self.assertIn("toolForm.Height = y + 26", helpers)
        self.assertIn('PositionControl designer, "cmdPreview", previewLeft, y, previewW, 42', installer)
        self.assertIn('PositionControl designer, "cmdPreview", MARGIN, y, contentW, BTN_H + 6', installer)

    def test_reset_stays_right_aligned_on_title_band(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn(
            "Private Sub LayoutGuidedTitleReset(ByVal toolForm As Object, ByVal M As Single, ByVal contentW As Single)",
            helpers,
        )
        self.assertNotIn("Dim titleAnchor As Object", helpers)
        self.assertNotIn("Set titleAnchor = GuidedTitleAnchor(toolForm)", helpers)
        self.assertIn("LayoutGuidedTitleReset toolForm, M, contentW", helpers)
        self.assertIn("resetTop = 8", helpers)
        self.assertIn("resetLeft = M + contentW - resetW", helpers)
        self.assertNotIn("GuidedTitleTextRight", helpers)
        self.assertNotIn("GuidedFallbackTitleTextRight", helpers)
        self.assertIn("Dim resetTop As Single", helpers)
        self.assertIn('toolForm.Controls("cmdReset").Move resetLeft, resetTop, resetW, resetH', helpers)
        self.assertNotIn("titleAnchor", helpers)
        self.assertNotIn("LayoutGuidedTitleReset toolForm, M, contentW, 8, introH", helpers)

        reset_call = helpers.index("LayoutGuidedTitleReset toolForm, M, contentW")
        flow_starts = helpers.index("    Dim introH As Single")
        self.assertLess(reset_call, flow_starts)
