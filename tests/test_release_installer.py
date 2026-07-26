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
