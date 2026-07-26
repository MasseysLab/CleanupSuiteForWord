from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_structural_safety_runner_syncs_runs_word_smoke_and_exports_fixture():
    source = (ROOT / "scripts" / "run_structural_safety_smoke.ps1").read_text(encoding="utf-8")

    assert "scripts\\sync_docm_code_only.ps1" in source
    assert "New-Object -ComObject Word.Application" in source
    assert "RunStructuralSafetySmokeCheck" in source
    assert '"modAlphaSmokeRunner.RunStructuralSafetySmokeCheck"' in source
    assert 'StartsWith("PASS|Structural Safety|")' in source
    assert "Regression - Structural Safety.docx" in source
    assert "[switch]$SkipFixtureExport" in source
    assert "if (-not $SkipFixtureExport)" in source
    assert "ExportAsFixedFormat($pdfPath, 17)" in source
    assert "structural_safety_smoke_" in source
