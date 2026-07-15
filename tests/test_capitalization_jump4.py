from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class CapitalizationJump4Tests(unittest.TestCase):
    def test_capitalization_has_editable_custom_exception_controls_and_storage(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        installer = read("src/installer/installer.bas")

        for snippet in [
            'cmdEditExceptions.Caption = "Edit custom exceptions"',
            "Private Sub cmdEditExceptions_Click()",
            "PromptForCapitalizationCustomExceptions",
            "GetCapitalizationCustomExceptions",
            "SetCapitalizationCustomExceptions",
            "CleanupSuiteCapitalizationExceptions",
            "InvalidateCapitalizationCache",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

        self.assertIn('"cmdEditExceptions"', installer)
        self.assertIn('Case "frmCapitalizationCleanup.cmdEditExceptions": ControlCaptionText = "Edit custom exceptions"', installer)

    def test_custom_exception_editor_uses_a_friendly_semicolon_list(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")

        for snippet in [
            '"Enter words exactly as you want them capitalized, separated by semicolons."',
            '"Example: MetaDataSuite; CleanUpSuite; JW"',
            "Private Function NormalizeCapitalizationExceptionList(",
            "exceptionText = NormalizeCapitalizationExceptionList(exceptionText)",
            "NormalizeCapitalizationExceptionList(GetCapitalizationCustomExceptions())",
            "Public Function SmokeFriendlyCapitalizationExceptions(",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

        self.assertNotIn("Enter custom capitalization exceptions as Name=Name", form)
        self.assertIn('" MetaDataSuite ; CleanUpSuite "', smoke)
        self.assertIn('"MetaDataSuite; CleanUpSuite"', smoke)
        self.assertNotIn("LoadCapitalizationCustomExceptionPairs", form)
        custom_loader = form[form.index("Private Sub LoadCapitalizationCustomExceptions("):]
        self.assertNotIn('Replace(normalized, "|", "=")', custom_loader)
        self.assertNotIn('separatorPos = InStr(normalized, "=")', custom_loader)

    def test_capitalization_uses_compiled_data_packs_and_document_cache(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        for snippet in [
            "Private mCapitalizationCacheKey As String",
            "Private mCapitalizationAbbreviationDict As Object",
            "Private Sub PrepareCapitalizationLookups()",
            "Private Function CapitalizationCacheKey() As String",
            "LoadCapitalizationCompiledDataPacks",
            "LoadCapitalizationCustomExceptions",
            '"appt."',
            '"LLC|LLC"',
            '"onedrive|OneDrive"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

    def test_capitalization_uses_word_heading_styles_when_repairing_paragraphs(self):
        form = read("src/forms/frmCapitalizationCleanup.bas")

        for snippet in [
            "Dim styleHeadingLike As Boolean",
            "styleHeadingLike = ParagraphUsesHeadingStyle(p)",
            "SmartRepairParagraph(original, useAbbreviations, protectAcronyms, protectNames, useHeadingHeuristics, titleCaseParentheticalText, useContextRules, styleHeadingLike)",
            "Private Function ParagraphUsesHeadingStyle(ByVal p As Paragraph) As Boolean",
            'styleName = LCase$(CStr(p.Style.NameLocal))',
            'Left$(styleName, 7) = "heading"',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, form)

    def test_public_versioning_reserves_095_for_beta(self):
        versioning = read("VERSIONING.md")

        for snippet in [
            "`0.9.0-alpha` | Official Alpha",
            "`0.9.5-beta` | Official Beta",
            "reserve the entire `0.9.5` milestone for Beta",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, versioning)


if __name__ == "__main__":
    unittest.main()
