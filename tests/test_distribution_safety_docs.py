from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def test_readme_exposes_free_signing_policy_privacy_and_unsigned_safety():
    readme = read("README.md")

    assert "## Code signing policy" in readme
    assert "Free code signing provided by" in readme
    assert "SignPath Foundation" in readme
    assert "(CODE_SIGNING_POLICY.md)" in readme
    assert "(PRIVACY.md)" in readme
    assert "(docs/Unsigned_Installation_Guide.md)" in readme


def test_unsigned_install_guide_preserves_windows_and_word_protections():
    guide = read("docs/Unsigned_Installation_Guide.md")

    assert "Get-FileHash" in guide
    assert "64-character result" in guide
    assert "Scan with" in guide
    assert "Microsoft Defender" in guide
    assert "**More info**" in guide
    assert "**Run anyway**" in guide
    assert "Do not disable or bypass that policy" in guide
    assert "turn off Microsoft Defender or SmartScreen" in guide
    assert "weaken Word macro-security settings" in guide
    assert "Do not use this template-only method for a hybrid release" in guide


def test_privacy_policy_describes_only_actual_network_and_local_processing():
    policy = read("PRIVACY.md")

    assert "does not provide accounts, advertising, analytics, telemetry" in policy
    assert "owner-only" in policy
    assert "does not transmit document content" in policy
    assert "at most once every seven days" in policy
    assert "GitHub-hosted release manifest" in policy
    assert "Periodic checks can be disabled" in policy


def test_signing_policy_keeps_manual_approval_and_unsigned_labeling():
    policy = read("CODE_SIGNING_POLICY.md")

    assert "Free code signing provided by" in policy
    assert "certificate by" in policy
    assert "receive explicit signing approval" in policy
    assert "multi-factor authentication" in policy
    assert "An unsigned build must be clearly identified as unsigned" in policy
