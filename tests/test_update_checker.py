from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def test_update_checker_is_packaged_and_uses_stable_manifest():
    manifest = read("src/manifest.txt")
    installer = read("src/installer/installer.bas")
    sync_script = read("scripts/sync_docm_code_only.ps1")
    checker = read("src/modules/modUpdateChecker.bas")

    assert "BUILDER  modules/modUpdateChecker.bas  modUpdateChecker_Code" in manifest
    assert 'CreateOrReplaceModule vbProj, "modUpdateChecker"' in installer
    assert '"modUpdateChecker")' in installer
    assert 'ComponentName = "modUpdateChecker"' in sync_script
    assert '"lblAutoSaveWarning", "chkUpdateChecks", "cmdCheckUpdates"' in sync_script
    assert "main/release/latest-release.ini" in checker
    assert 'UPDATE_ALLOWED_URL_PREFIX As String = "https://github.com/MasseysLab/CleanupSuiteForWord/"' in checker


def test_update_checks_are_opt_in_visible_and_reversible():
    launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
    helpers = read("src/modules/modCleanupHelpers.bas")
    checker = read("src/modules/modUpdateChecker.bas")
    smoke = read("src/modules/modAlphaSmokeRunner.bas")

    assert 'chkUpdateChecks.Caption = "Check periodically for updates"' in launcher
    assert "chkUpdateChecks.Value = GetCleanupSuiteUpdateChecksEnabled()" in launcher
    assert "SetCleanupSuiteUpdateChecksEnabled chkUpdateChecks.Value" in launcher
    assert "ExplainCleanupSuiteUpdateChecksDisabled" in launcher
    assert "CheckForCleanupSuiteUpdates True" in launcher
    assert "SetCleanupSuiteUpdateChecksEnabled True" in helpers
    assert "To turn them back on" in checker
    assert "You can still choose Check for Updates at any time." in checker
    assert 'SmokeRequireControl launcher, "chkUpdateChecks"' in smoke
    assert 'SmokeRequireControl launcher, "cmdCheckUpdates"' in smoke


def test_update_check_is_periodic_quiet_and_does_not_mutate_documents():
    checker = read("src/modules/modUpdateChecker.bas")
    launcher_module = read("src/modules/modCleanupLauncher.bas")

    assert "Private Const UPDATE_INTERVAL_DAYS As Long = 7" in checker
    assert "If Not CleanupSuiteUpdateCheckIsDue() Then Exit Sub" in checker
    assert "CheckForCleanupSuiteUpdates False" in checker
    assert "If manualCheck Then" in checker
    assert "ActiveDocument." not in checker
    assert "Public Sub AutoExec()" in launcher_module
    assert "ScheduleCleanupSuiteUpdateCheck" in launcher_module
