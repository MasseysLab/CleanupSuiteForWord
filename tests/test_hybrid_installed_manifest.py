import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE = (ROOT / "src/modules/modHybridBridge.bas").read_text(encoding="utf-8")
ENGINE_PATH_POLICY = (
    ROOT / "engine/CleanupSuite.Engine/JobPathPolicy.cs"
).read_text(encoding="utf-8")
SMOKE = (
    ROOT / "scripts/run_hybrid_installed_manifest_smoke.ps1"
).read_text(encoding="utf-8")


class HybridInstalledManifestTests(unittest.TestCase):
    def test_runtime_prefers_matched_install_unless_dev_override_is_explicit(self):
        self.assertIn("Private Function HybridResolveTrustedEngine(", BRIDGE)
        self.assertIn('Environ$("CLEANUPSUITE_HYBRID_DEV_ENGINE")', BRIDGE)
        self.assertIn("HybridResolveTrustedDevelopmentEngine(", BRIDGE)
        self.assertIn("HybridResolveMatchedInstalledEngine(", BRIDGE)
        self.assertIn(
            "If Not HybridResolveTrustedEngine(enginePath, expectedEngineHash, failureMessage) Then Exit Function",
            BRIDGE,
        )

    def test_installed_manifest_binds_all_required_components(self):
        for snippet in [
            r"\MasseysLab\CleanupSuiteForWord",
            r"\installation-manifest.json",
            '"manifestVersion|suiteVersion|protocolVersion|publisher|authenticodeRequired|components"',
            '"word-template"',
            '"analysis-engine"',
            '"tool-definitions"',
            '"rules"',
            r'"Engine\CleanupSuite.Engine.exe"',
            r'"Contracts\Hybrid\v1\protocol.json"',
            r'"Contracts\Hybrid\v1\operation-vocabulary.json"',
            "FileLen(componentPath) <> componentLength",
            'HybridSha256File(componentPath, componentId = "word-template") <> componentHash',
            "Optional ByVal allowSharedRead As Boolean = False",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, BRIDGE)

    def test_mismatch_fails_closed_and_recommends_repair(self):
        self.assertIn(
            "CleanupSuite's matched analysis engine is unavailable or changed.",
            BRIDGE,
        )
        self.assertIn(
            "Run CleanupSuite Setup and choose Repair/Reinstall.",
            BRIDGE,
        )
        self.assertIn(
            "If Not HybridValidateCapabilityHandshake(enginePath, failureMessage) Then Exit Function",
            BRIDGE,
        )

    def test_installed_manifest_smoke_uses_one_explicit_test_jobs_root(self):
        self.assertIn("CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT", ENGINE_PATH_POLICY)
        self.assertIn("Path.GetFullPath(testRoot)", ENGINE_PATH_POLICY)
        self.assertIn('Environ$("CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT")', BRIDGE)
        self.assertIn("$env:CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT", SMOKE)
        self.assertIn("RunHybridInstalledManifestSmokeCheck", SMOKE)
        self.assertIn("CLEANUPSUITE_HYBRID_DEV_ENGINE", SMOKE)


if __name__ == "__main__":
    unittest.main()
