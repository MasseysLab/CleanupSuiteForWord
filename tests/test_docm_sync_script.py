from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class DocmSyncScriptTests(unittest.TestCase):
    def test_sync_script_includes_runtime_critical_components(self):
        script = (ROOT / "scripts" / "sync_docm_code_only.ps1").read_text(encoding="utf-8")

        self.assertIn('ComponentName = "modCleanupLauncher"', script)
        self.assertIn('ComponentName = "frmPreviewActions"', script)
        self.assertIn('ComponentName = "frmCleanupSuiteLauncher"', script)


if __name__ == "__main__":
    unittest.main()
