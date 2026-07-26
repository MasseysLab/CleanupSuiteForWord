import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE = (ROOT / "src/modules/modHybridBridge.bas").read_text(encoding="utf-8")
JSON_MODULE = (ROOT / "src/modules/modHybridJson.bas").read_text(encoding="utf-8")
UNICODE_FORM = (ROOT / "src/forms/frmUnicodeCleanup.bas").read_text(
    encoding="utf-8"
)
INSTALLER = (ROOT / "src/installer/installer.bas").read_text(encoding="utf-8")
MANIFEST = (ROOT / "src/manifest.txt").read_text(encoding="utf-8")
SYNC = (ROOT / "scripts/sync_docm_code_only.ps1").read_text(encoding="utf-8")


class HybridUnicodePilotTests(unittest.TestCase):
    def test_bridge_uses_direct_process_and_exact_local_trust(self):
        self.assertIn("CreateProcessW", BRIDGE)
        self.assertIn("BCryptOpenAlgorithmProvider", BRIDGE)
        self.assertIn("CLEANUPSUITE_HYBRID_DEV_ENGINE", BRIDGE)
        self.assertIn("CLEANUPSUITE_HYBRID_DEV_SHA256", BRIDGE)
        self.assertIn("HybridSha256File(enginePath) <> expectedHash", BRIDGE)
        self.assertNotIn("WScript.Shell", BRIDGE)
        self.assertNotIn('Shell "', BRIDGE)
        self.assertNotIn("PowerShell", BRIDGE)

    def test_capability_handshake_is_allow_list_validated(self):
        self.assertIn('HybridCaptureEngineOutput(enginePath, "--capabilities", 10)', BRIDGE)
        self.assertIn("HybridRequireExactKeys root", BRIDGE)
        self.assertIn('"engine-capabilities"', BRIDGE)
        self.assertIn('"invisible-unicode-cleaner"', BRIDGE)
        self.assertIn('"selected-characters"', BRIDGE)
        self.assertIn('"replaceText"', BRIDGE)
        self.assertIn('"authenticodeRequiredForOfficialBeta"', BRIDGE)

    def test_json_parser_never_evaluates_script(self):
        self.assertIn("HybridJsonParseObject", JSON_MODULE)
        self.assertIn("duplicate key", JSON_MODULE)
        self.assertNotIn("ScriptControl", JSON_MODULE)
        self.assertNotIn("Eval(", JSON_MODULE)
        self.assertNotIn("Execute(", JSON_MODULE)

    def test_preview_uses_engine_candidates_and_minimal_highlights(self):
        preview_start = UNICODE_FORM.index("If previewOnly Then")
        apply_start = UNICODE_FORM.index(
            "If mHybridAnalysis Is Nothing", preview_start
        )
        preview = UNICODE_FORM[preview_start:apply_start]
        self.assertIn("HybridAnalyzeInvisibleUnicode", preview)
        self.assertIn("HybridCandidateAbsoluteRange", preview)
        self.assertIn("ApplyPreviewHighlight hybridRange", preview)
        self.assertIn("UnicodeHybridNeedsAdjacentFallback", preview)
        self.assertNotIn("CountPreviewFindMatches", preview)
        self.assertNotIn("HighlightUnicodeMatches targetRange", preview)

    def test_apply_revalidates_everything_before_undo_or_mutation(self):
        revalidate = UNICODE_FORM.index("HybridRevalidateInvisibleUnicode")
        mark_start = UNICODE_FORM.index('MarkCleanupStart "Unicode Cleaner"', revalidate)
        undo_start = UNICODE_FORM.index(
            'undoRec.StartCustomRecord "Cleanup Suite - Unicode Cleaner"',
            mark_start,
        )
        mutation = UNICODE_FORM.index(
            "applyRange.Text = HybridCandidateReplacement", undo_start
        )
        self.assertLess(revalidate, mark_start)
        self.assertLess(mark_start, undo_start)
        self.assertLess(undo_start, mutation)
        self.assertIn(
            "For idx = applyCandidates.Count To 1 Step -1",
            UNICODE_FORM[undo_start:mutation],
        )

    def test_bridge_modules_are_installed_assembled_and_synced(self):
        for module in ("modHybridJson", "modHybridBridge"):
            self.assertIn(
                f'CreateOrReplaceModule vbProj, "{module}"',
                INSTALLER,
            )
            self.assertIn(f'"{module}"', INSTALLER)
            self.assertIn(f"modules/{module}.bas", MANIFEST)
            self.assertIn(f"src\\modules\\{module}.bas", SYNC)


if __name__ == "__main__":
    unittest.main()
