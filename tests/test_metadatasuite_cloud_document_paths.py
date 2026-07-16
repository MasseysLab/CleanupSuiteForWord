from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "src/modules/modMetaDataSuite.bas").read_text(
    encoding="utf-8", errors="replace"
)


def test_personal_onedrive_urls_resolve_to_an_existing_local_package():
    assert 'PERSONAL_URL_PREFIX As String = "https://d.docs.live.net/"' in MODULE
    assert 'Environ$("OneDriveConsumer")' in MODULE
    assert 'Environ$("OneDrive")' in MODULE
    assert 'repairedPath = Replace(candidatePath, "^", ".")' in MODULE
    assert "If FileExists(repairedPath) Then" in MODULE


def test_dashboard_uses_the_resolved_physical_path_instead_of_word_cloud_urls():
    assert "resolvedPath = GetDocumentPhysicalPath(doc)" in MODULE
    assert "ResolvePreferredDocxTarget = resolvedPath" in MODULE
    assert "PeekPreferredDocxTarget = resolvedPath" in MODULE
    assert "ResolvePreferredDocxTarget = doc.FullName" not in MODULE
    assert "PeekPreferredDocxTarget = doc.FullName" not in MODULE


def test_open_document_matching_uses_the_same_physical_path_resolution():
    assert MODULE.count("If DocumentMatchesPhysicalPath(doc, filePath) Then") >= 2
    assert "documentPath = GetDocumentPhysicalPath(doc)" in MODULE
    assert "targetPath = ResolveExistingPathVariant(filePath)" in MODULE
