from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "Testing Files" / "Regression - Structural Safety.docx"
PRACTICE_FIXTURE = (
    ROOT
    / "Practice - Try CleanupSuite Here"
    / "Testing Files"
    / "Regression - Structural Safety.docx"
)


def package_text(path: Path, member: str) -> str:
    with ZipFile(path) as package:
        return package.read(member).decode("utf-8", errors="ignore")


def test_structural_safety_fixture_is_mirrored_and_deterministic():
    assert FIXTURE.is_file()
    assert PRACTICE_FIXTURE.is_file()
    assert sha256(FIXTURE.read_bytes()).digest() == sha256(PRACTICE_FIXTURE.read_bytes()).digest()


def test_structural_safety_fixture_contains_real_protected_word_structures():
    document_xml = package_text(FIXTURE, "word/document.xml")

    assert "CleanupSuite Structural Safety Regression" in document_xml
    assert "CleanupSuiteProtectedBlank" in document_xml
    assert "CleanupSuiteStructuralSafety" in document_xml
    assert "PROTECTED HIDDEN CONTENT" in document_xml
    assert "FINAL CONTENT SENTINEL" in document_xml
    assert "\u00a0" in document_xml
    assert "<w:bookmarkStart" in document_xml
    assert "<w:fldSimple" in document_xml
    assert "<w:sdt" in document_xml
    assert "<w:vanish" in document_xml
    assert "<w:tblHeader" in document_xml
    assert "<w:gridSpan" in document_xml
    assert document_xml.count("<w:tbl>") >= 9


def test_structural_safety_fixture_uses_fixed_table_geometry():
    document_xml = package_text(FIXTURE, "word/document.xml")

    assert document_xml.count('<w:tblW w:type="dxa" w:w="9360"') >= 8
    assert document_xml.count('<w:tblInd w:w="120" w:type="dxa"') >= 9
    assert "w:type=\"pct\"" not in document_xml
