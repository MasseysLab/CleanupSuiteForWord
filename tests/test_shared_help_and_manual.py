from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class SharedHelpAndManualTests(unittest.TestCase):
    TOOL_HELP_KEYS = {
        "frmUnicodeCleanup.bas": "Unicode",
        "frmPunctuationCleanup.bas": "Punctuation",
        "frmSpacingCleanup.bas": "Spacing",
        "frmCapitalizationCleanup.bas": "Capitalization",
        "frmListCleanup.bas": "List",
        "frmParagraphCleanup.bas": "Paragraph",
        "frmDuplicateDetector.bas": "Duplicate",
        "frmFontNormalizer.bas": "Font",
        "frmTableCleaner.bas": "Table",
        "frmBreakNormalizer.bas": "Break",
        "frmDocumentTrim.bas": "DocumentTrim",
        "frmFormattingStripper.bas": "Formatting",
        "frmHyperlinkRemover.bas": "Hyperlink",
        "frmSoftReturnConverter.bas": "SoftReturn",
        "frmMetaDataSuite.bas": "MetaDataSuite",
        "frmFinalReview.bas": "FinalReview",
        "frmStyleCleanup.bas": "Style",
        "frmFootnoteRemover.bas": "Footnote",
        "frmHeaderFooterStandardizer.bas": "HeaderFooter",
        "frmObjectRemover.bas": "Object",
    }

    LAUNCHER_HELP_KEYS = {
        "cmdHelpUnicode": "Unicode",
        "cmdHelpPunct": "Punctuation",
        "cmdHelpSpacing": "Spacing",
        "cmdHelpCap": "Capitalization",
        "cmdHelpList": "List",
        "cmdHelpPara": "Paragraph",
        "cmdHelpDuplicate": "Duplicate",
        "cmdHelpFont": "Font",
        "cmdHelpTable": "Table",
        "cmdHelpBreak": "Break",
        "cmdHelpTrim": "DocumentTrim",
        "cmdHelpFormat": "Formatting",
        "cmdHelpHyperlink": "Hyperlink",
        "cmdHelpSoftReturn": "SoftReturn",
        "cmdHelpMetadata": "MetaDataSuite",
        "cmdHelpFinalReview": "FinalReview",
        "cmdHelpStyle": "Style",
        "cmdHelpFootnote": "Footnote",
        "cmdHelpHeaderFooter": "HeaderFooter",
        "cmdHelpObject": "Object",
    }

    def test_tool_forms_do_not_have_help_buttons_or_handlers(self):
        for filename, key in self.TOOL_HELP_KEYS.items():
            source = read(f"src/forms/{filename}")
            with self.subTest(form=filename):
                self.assertNotIn(f'ShowCleanupToolHelp "{key}"', source)
                self.assertNotIn("cmdHelp_Click", source)
                self.assertNotIn("MsgBox h, vbInformation", source)
                self.assertNotIn("Help  --", source)

    def test_launcher_help_buttons_use_the_same_shared_help_text(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        for handler, key in self.LAUNCHER_HELP_KEYS.items():
            with self.subTest(handler=handler):
                pattern = (
                    r"Private Sub "
                    + re.escape(handler)
                    + r'_Click\(\)\s+ShowCleanupToolHelp "'
                    + re.escape(key)
                    + r'"\s+End Sub'
                )
                self.assertRegex(launcher, pattern)

        self.assertNotIn("capitalisation", launcher.lower())
        self.assertNotIn("behaviour", launcher.lower())
        self.assertNotIn("Help  --", launcher)

    def test_shared_help_text_is_professional_and_contains_every_tool(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Public Sub ShowCleanupToolHelp(ByVal toolKey As String)", helpers)
        self.assertIn("Public Function CleanupHelpText(ByVal toolKey As String) As String", helpers)
        self.assertIn("Public Function CleanupHelpTitle(ByVal toolKey As String) As String", helpers)

        for key in self.TOOL_HELP_KEYS.values():
            with self.subTest(tool_key=key):
                self.assertIn(f'Case "{key}"', helpers)

        self.assertNotIn("capitalisation", helpers.lower())
        self.assertNotIn("behaviour", helpers.lower())
        self.assertNotIn("Help  --", helpers)
        self.assertNotIn("  --", helpers)

    def test_capitalization_help_has_one_source_from_main_menu_only(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        form = read("src/forms/frmCapitalizationCleanup.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn('ShowCleanupToolHelp "Capitalization"', launcher)
        self.assertNotIn('ShowCleanupToolHelp "Capitalization"', form)
        self.assertNotIn("cmdHelp_Click", form)
        self.assertEqual(2, helpers.count('Case "Capitalization"'))
        self.assertIn("Recommended repair", helpers)
        self.assertNotIn("All (Smart)", helpers)

    def test_launcher_has_user_manual_pdf_button(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("Private Sub cmdUserManual_Click()", launcher)
        self.assertIn("OpenCleanupUserManualPdf", launcher)
        self.assertIn("Public Sub OpenCleanupUserManualPdf()", helpers)
        self.assertIn("CLEANUP_SUITE_USER_MANUAL_PDF_URL", helpers)
        self.assertIn(".pdf", helpers)
        self.assertIn("cmdUserManual", installer)
        self.assertIn('Case "cmdUserManual": cap = "User Manual"', installer)
        self.assertIn('PositionControl designer, "cmdUserManual"', installer)

    def test_user_manual_pdf_artifact_exists(self):
        pdf_path = ROOT / "documents" / "CleanupSuite_User_Manual.pdf"
        self.assertTrue(pdf_path.exists())
        self.assertGreater(pdf_path.stat().st_size, 10_000)


if __name__ == "__main__":
    unittest.main()
