from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class RefreshStartupTemplateScriptTests(unittest.TestCase):
    def test_refresh_script_uses_repo_relative_paths(self):
        script = (ROOT / "scripts" / "refresh_startup_dotm.ps1").read_text(encoding="utf-8")

        self.assertIn('$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path', script)
        self.assertIn('$StartupDir = Join-Path $env:APPDATA "Microsoft\\Word\\STARTUP"', script)
        self.assertIn('"CleanupSuite external backups\\startup-dotm"', script)
        self.assertNotIn('$RepoRoot "backups\\startup-dotm"', script)
        self.assertNotIn('C:\\Users\\Chris\\PROGRAMMING\\VBA\\CleanupSuite for MS Word', script)

    def test_refresh_script_uses_backup_temp_and_openxml_copy_pattern(self):
        script = (ROOT / "scripts" / "refresh_startup_dotm.ps1").read_text(encoding="utf-8")

        for snippet in [
            'Join-Path $StartupDir "CleanupSuite.dotm"',
            'Join-Path $BackupDir ("CleanupSuite.startup-backup-',
            'Convert-DocmPackageToDotmPackage',
            'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
            'Set-ZipEntryText -ZipPath $TargetDotmPath -EntryName "[Content_Types].xml"',
            'Copy-Item -LiteralPath $tempDotmPath -Destination $startupDotmPath -Force',
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, script)

    def test_refresh_script_does_not_use_word_automation(self):
        script = (ROOT / "scripts" / "refresh_startup_dotm.ps1").read_text(encoding="utf-8")

        self.assertNotIn("New-Object -ComObject Word.Application", script)
        self.assertNotIn("SaveAs2(", script)
        self.assertNotIn("WINWORD", script)


if __name__ == "__main__":
    unittest.main()
