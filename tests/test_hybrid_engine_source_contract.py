import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ENGINE_ROOT = ROOT / "engine"
ENGINE_SOURCE = ENGINE_ROOT / "CleanupSuite.Engine"


class HybridEngineSourceContractTests(unittest.TestCase):
    def test_engine_has_no_external_package_dependency(self):
        projects = list(ENGINE_ROOT.rglob("*.csproj"))
        self.assertGreaterEqual(len(projects), 2)
        for project in projects:
            text = project.read_text(encoding="utf-8")
            self.assertNotIn("<PackageReference", text)

    def test_engine_source_has_no_word_network_shell_or_elevation_surface(self):
        source = "\n".join(
            path.read_text(encoding="utf-8")
            for path in sorted(ENGINE_SOURCE.glob("*.cs"))
        )
        forbidden = (
            "Microsoft.Office.Interop",
            "Word.Application",
            "System.Net",
            "HttpClient",
            "WebClient",
            "Process.Start",
            "ShellExecute",
            "cmd.exe",
            "powershell.exe",
            "runas",
        )
        for token in forbidden:
            with self.subTest(token=token):
                self.assertNotIn(token, source)

    def test_engine_exposes_only_fixture_and_unicode_pilot_analyzers(self):
        constants = (
            ENGINE_SOURCE / "ContractConstants.cs"
        ).read_text(encoding="utf-8")
        capabilities = (
            ENGINE_SOURCE / "Capabilities.cs"
        ).read_text(encoding="utf-8")
        self.assertIn(
            'FixtureToolId = "contract-fixture"',
            constants,
        )
        self.assertIn(
            'FixtureAnalysisMode = "replace-literal"',
            constants,
        )
        self.assertIn("ContractConstants.FixtureToolId", capabilities)
        self.assertIn(
            'UnicodeToolId = "invisible-unicode-cleaner"',
            constants,
        )
        self.assertIn("ContractConstants.UnicodeToolId", capabilities)
        self.assertNotIn("punctuation-normalizer", capabilities)
        self.assertNotIn("duplicate-paragraph", capabilities)

    def test_engine_security_claims_match_protocol(self):
        protocol = json.loads(
            (
                ROOT / "contracts" / "hybrid" / "v1" / "protocol.json"
            ).read_text(encoding="utf-8")
        )
        source = (
            ENGINE_SOURCE / "Capabilities.cs"
        ).read_text(encoding="utf-8")
        self.assertFalse(protocol["applySafety"]["engineMayEditWord"])
        for claim in (
            "editsWordDocuments = false",
            "requiresNetwork = false",
            "opensListeningEndpoint = false",
            "runsAsService = false",
            "requiresElevation = false",
            "logsDocumentContent = false",
        ):
            self.assertIn(claim, source)

    def test_ci_runs_isolated_engine_gate(self):
        workflow = (
            ROOT / ".github" / "workflows" / "ci.yml"
        ).read_text(encoding="utf-8")
        runner = (
            ROOT / "scripts" / "run_hybrid_engine_tests.ps1"
        ).read_text(encoding="utf-8")
        self.assertIn("actions/setup-dotnet@v4", workflow)
        self.assertIn("run_hybrid_engine_tests.ps1", workflow)
        self.assertIn("CleanupSuite.Engine.ContractTests.csproj", runner)
        self.assertIn("validate_hybrid_engine_capabilities.py", runner)
        self.assertIn("validate_hybrid_engine_result.py", runner)

    def test_development_publish_is_self_contained_but_not_committed(self):
        publisher = (
            ROOT / "scripts" / "publish_hybrid_engine_dev.ps1"
        ).read_text(encoding="utf-8")
        gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")
        self.assertIn("--self-contained true", publisher)
        self.assertIn("-p:PublishSingleFile=true", publisher)
        self.assertIn("-p:PublishTrimmed=false", publisher)
        self.assertIn("validate_hybrid_engine_capabilities.py", publisher)
        self.assertIn("/build/", gitignore)


if __name__ == "__main__":
    unittest.main()
