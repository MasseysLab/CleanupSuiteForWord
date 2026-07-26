from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class WordAlphaSmokeRunnerTests(unittest.TestCase):
    def test_capitalization_has_a_focused_runtime_smoke_entry_point(self):
        source = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn("Public Function RunCapitalizationAlphaSmokeCheck() As String", source)
        self.assertIn("RunCapitalizationAlphaSmokeCheck = CStr(SmokeCapitalizationCheck())", source)

    def test_checklist_names_the_required_smoke_targets(self):
        checklist = read("docs/Alpha_Readiness_Checklist_0.8.5.md")
        for phrase in [
            "open the main launcher",
            "open Document Trim as the simplest guided form",
            "open Invisible Unicode Cleaner as a Custom-mode form",
            "open Capitalization Fixer as the closest working reference form",
            "open Header / Footer Standardizer as the special-case layout form",
            "run at least one Preview flow",
        ]:
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, checklist)

    def test_runner_and_vba_harness_are_wired_to_named_forms_and_preview(self):
        script = read("scripts/run_word_alpha_smoke.ps1")
        smoke = read("src/modules/modAlphaSmokeRunner.bas")
        for snippet in [
            "frmCleanupSuiteLauncher",
            "frmDocumentTrim",
            "frmUnicodeCleanup",
            "frmCapitalizationCleanup",
            "cmdEditExceptions",
            "frmHeaderFooterStandardizer",
            "frmObjectRemover",
            "cmdRiskPlacementObjectTables",
            "cmdRiskPlacementHeaderFooterClear",
            "frmPreviewActions",
            "RunAlphaSmokeChecklist",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, script + smoke)

    def test_smoke_module_exposes_result_writer_and_named_checks(self):
        smoke = read("src/modules/modAlphaSmokeRunner.bas")
        for snippet in [
            "Public Function RunAlphaSmokeChecklist(",
            "Private Function SmokeLauncherCheck(",
            "Private Function SmokeDocumentTrimCheck(",
            "Private Function SmokeUnicodeCheck(",
            "Private Function SmokeCapitalizationCheck(",
            "Private Function SmokeHeaderFooterCheck(",
            "Private Function SmokeFormattingCleanerCheck(",
            "Private Function SmokeStyleCleanupApplyCheck(",
            "Private Function SmokeRiskControlsCheck(",
            "Private Function SmokePreviewFlowCheck(",
            "Private Function SmokePreviewSummaryTableCheck(",
            "Private Sub SmokeRequireSummaryRowFits(",
            "Panel Preview OFF did not update its state label.",
            "Panel Preview OFF still exposed the unstable direct re-preview shortcut.",
            "Panel Preview OFF left navigation visible.",
            'panel.ConfigureSummary toolForm, "Document Trim", flowRows',
            'panel.ConfigureSummary toolForm, "Object Remover", wideRows',
            'toolForm.RunAfterPreview',
            'TTF-19_StyleCleanup.docx',
            "PASS|",
            "FAIL|",
            "WARN|",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, smoke)

        self.assertNotIn('CallByName toolForm, "PreviewFromPanel", VbMethod', smoke)

    def test_powershell_runner_calls_word_and_writes_a_report(self):
        script = read("scripts/run_word_alpha_smoke.ps1")
        for snippet in [
            "New-Object -ComObject Word.Application",
            "Practice - Try CleanupSuite Here\\CleanupSuite.docm",
            "RunAlphaSmokeChecklist",
            "alpha_smoke_report",
            "PASS|Launcher",
            "PASS|Formatting Cleaner",
            "PASS|Risk Controls",
            "PASS|Preview Summary Table",
            '"Item"',
            '"#"',
            '" (unhighlighted)"',
            "SUMMARY_BOTTOM_GAP",
            "If Me.InsideHeight < desiredInsideH Then",
            "ShowPreviewActionsSummary",
            "WARN|Visual review",
            "p.Range.Font.Reset",
            "p.Range.ParagraphFormat.Reset",
            "ClearCharacterDirectFormatting",
            "ClearParagraphDirectFormatting",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, script)

    def test_main_check_script_runs_the_regression_suite(self):
        script = read("check.ps1")
        self.assertIn("python -m unittest discover -s tests", script)
        self.assertIn("Regression tests failed.", script)

    def test_docs_point_release_workflow_to_the_smoke_runner(self):
        readme = read("README.md")
        for snippet in [
            "run_word_alpha_smoke.ps1",
            "Run the Word alpha smoke runner",
            "visual review",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, readme)


if __name__ == "__main__":
    unittest.main()
