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
    assert values["version"] == "0.9.0-alpha"
    assert values["tag"] == "v0.9.0-alpha"
    assert values["template_sha256"] == hashlib.sha256(template.read_bytes()).hexdigest()
    assert values["setup_url"].endswith("/main/release/CleanupSuiteForWord-Setup.exe")


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
    assert 'Require(values, "template_sha256")' in source
    assert "SHA256.Create()" in source
    assert 'GetManifestResourceStream("CleanupSuite.dotm")' in source
    assert '"/resource:$releaseTemplate,CleanupSuite.dotm"' in build
    assert "MasseysLab" in source
    assert "CleanupSuiteForWord-Setup.exe" in build
