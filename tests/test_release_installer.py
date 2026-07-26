import hashlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def manifest_values() -> dict[str, str]:
    values: dict[str, str] = {}
    for raw in read("release/latest-release.ini").splitlines():
        line = raw.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip()
    return values


def test_stable_installer_pointer_and_special_vba_template_are_consistent():
    setup = ROOT / "release" / "CleanupSuiteForWord-Setup.exe"
    template = ROOT / "release" / "CleanupSuite.dotm"
    special_release_notes = read("docs/Release_Notes_v0.9.3-beta-vba-final.md")
    values = manifest_values()

    assert setup.stat().st_size > 100_000
    assert template.stat().st_size > 100_000
    assert values["version"] == "0.9.2-alpha"
    assert values["tag"] == "v0.9.2-alpha"
    assert values["template_sha256"] == (
        "a256a27bb58583c163687c45dc70dcb9074212d73a85d027b33521be48cf9cc0"
    )
    assert values["setup_url"].endswith("/main/release/CleanupSuiteForWord-Setup.exe")
    assert (
        hashlib.sha256(template.read_bytes()).hexdigest().upper()
        in special_release_notes
    )


def test_installer_is_per_user_bounded_and_does_not_weaken_word():
    source = read("installer/CleanupSuiteForWord.Setup.cs")

    assert "Environment.SpecialFolder.ApplicationData" in source
    assert '"Microsoft", "Word", "STARTUP"' in source
    assert "Registry.CurrentUser" in source
    assert "Registry.LocalMachine" not in source
    assert "Program.BackupLimit" in source
    assert "internal const int BackupLimit = 3" in source
    assert 'Process.GetProcessesByName("WINWORD")' in source
    assert "Normal.dotm" not in source
    assert "Trusted Location" not in source
    assert "VBAWarnings" not in source


def test_beta_installer_embeds_one_matched_offline_package():
    source = read("installer/CleanupSuiteForWord.Setup.cs")
    build = read("installer/build-installer.ps1")

    assert 'DisplayVersion = "0.9.5-beta"' in source
    assert 'GetManifestResourceStream(name)' in source
    assert '"CleanupSuite.Payload.Manifest.json"' in source
    assert '"CleanupSuite.Payload.Template.dotm"' in source
    assert '"CleanupSuite.Payload.Engine.exe"' in source
    assert '"CleanupSuite.Payload.Protocol.json"' in source
    assert '"CleanupSuite.Payload.Operations.json"' in source
    assert "StageAndVerify" in source
    assert "SHA256.Create()" in source
    assert "installation-manifest.json" in source
    assert "WinVerifyTrust" in source
    assert "AuthenticodeTrust.IsTrusted(stagedEngine)" in source
    assert "AuthenticodeTrust.IsTrusted(Application.ExecutablePath)" in source
    assert "PublishRelease" in build
    assert "SigningCertificateThumbprint" in build
    assert "Official hybrid release builds require -SigningCertificateThumbprint." in build
    assert "CleanupSuite.Payload.Engine.exe" in build
    assert "--self-contained true" in read("scripts/publish_hybrid_engine_dev.ps1")
    assert "MasseysLab" in source
    assert "CleanupSuiteForWord-Setup.exe" in build


def test_beta_installer_is_state_aware_and_maintenance_safe():
    source = read("installer/CleanupSuiteForWord.Setup.cs")

    for state in [
        "Absent",
        "Older",
        "CurrentHealthy",
        "CurrentDamaged",
        "Newer",
    ]:
        assert state in source
    for action in ["Install", "Update", "Repair", "Uninstall"]:
        assert action in source
    assert '"Repair/Reinstall"' in source
    assert "ValidateRequestedAction" in source
    assert "RollbackEntry.Capture" in source
    assert "RestoreRollback" in source
    assert "rollback.Add(RollbackEntry.Capture(paths.InstallerCopyPath))" in source
    assert "RetainNewestBackups(paths.BackupFolder, Program.BackupLimit)" in source
    assert "user settings will be kept" in source
    assert 'key.SetValue("ModifyPath"' in source
    assert 'key.SetValue("NoRepair", 0' in source
    assert "DeleteSubKeyTree(UninstallKeyPath" in source


def test_hybrid_installer_matrix_is_isolated_and_covers_all_states():
    script = read("scripts/run_hybrid_installer_matrix.ps1")

    assert "build\\installer-matrix-" in script
    assert "/test-root=" in script
    assert 'Assert-State $report "Absent"' in script
    assert 'Assert-State $report "CurrentHealthy"' in script
    assert 'Assert-State $report "CurrentDamaged"' in script
    assert 'Assert-State $report "Older"' in script
    assert 'Invoke-SetupAction "install"' in script
    assert 'Invoke-SetupAction "update"' in script
    assert 'Invoke-SetupAction "repair"' in script
    assert 'Invoke-SetupAction "uninstall"' in script
    assert "Expected exactly 3 retained backups" in script
    assert "Repair removed the user-settings sentinel" in script
    assert "Uninstall removed the user-settings sentinel" in script
