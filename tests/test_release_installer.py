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


def test_release_artifacts_and_manifest_are_consistent():
    setup = ROOT / "release" / "CleanupSuiteForWord-Setup.exe"
    template = ROOT / "release" / "CleanupSuite.dotm"
    values = manifest_values()

    assert setup.stat().st_size > 100_000
    assert template.stat().st_size > 100_000
    assert values["version"] == "0.9.3-beta-vba-final"
    assert values["tag"] == "v0.9.3-beta-vba-final"
    assert values["template_sha256"] == hashlib.sha256(template.read_bytes()).hexdigest()
    assert values["setup_sha256"] == hashlib.sha256(setup.read_bytes()).hexdigest()
    assert values["setup_url"].endswith("/main/release/CleanupSuiteForWord-Setup.exe")
    assert values["template_url"].endswith(
        "/v0.9.3-beta-vba-final/CleanupSuite.dotm"
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


def test_installer_uses_stable_manifest_hash_and_embedded_fallback():
    source = read("installer/CleanupSuiteForWord.Setup.cs")
    build = read("installer/build-installer.ps1")

    assert "Program.ManifestUrl" in source
    assert "RequestCacheLevel.NoCacheNoStore" in source
    assert 'Program.ManifestUrl + "?v=" + Uri.EscapeDataString(Program.DisplayVersion)' in source
    assert 'templateUrl += "?v=" + Uri.EscapeDataString(Program.DisplayVersion)' in source
    assert 'Require(values, "template_sha256")' in source
    assert "SHA256.Create()" in source
    assert 'GetManifestResourceStream("CleanupSuite.dotm")' in source
    assert '"/resource:$releaseTemplate,CleanupSuite.dotm"' in build
    assert "MasseysLab" in source
    assert "CleanupSuiteForWord-Setup.exe" in build


def test_default_vba_only_setup_is_stable_state_aware_and_hash_bound():
    source = read("installer/CleanupSuiteForWord.VbaOnly.Setup.cs")
    build = read("installer/build-vba-only-installer.ps1")
    matrix = read("scripts/run_vba_only_installer_matrix.ps1")

    assert 'DisplayVersion = "0.9.3-beta-vba-final"' in source
    assert "0.9.3 Stable - Final VBA-Only Release" in source
    assert "4f815ee0df4bacbda5150a82db2d66739592fa2098184c8d7ec7b535d73ab10c" in source
    assert "GetManifestResourceStream(\"CleanupSuite.dotm\")" in source
    assert "The installed template matches the published stable release." in source
    for action in ["Install", "Update", "Repair/Reinstall", "Uninstall"]:
        assert action in source
    assert "Registry.CurrentUser" in source
    assert "Registry.LocalMachine" not in source
    assert "Normal.dotm" not in source
    assert "VBAWarnings" not in source
    assert "v0.9.3-beta-vba-final" in build
    assert "release-rollback" in build
    assert "/test-root=" in matrix
    assert "VBA-Only Stable Version" in matrix
