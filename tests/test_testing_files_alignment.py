from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import unittest


ROOT = Path(__file__).resolve().parents[1]
ROOT_TESTS = ROOT / "Testing Files"
PRACTICE_TESTS = ROOT / "Practice - Try CleanupSuite Here" / "Testing Files"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


EXPECTED_FIXTURES = [
    ("TTF-01_InvisibleUnicodeCleaner.docx", "Invisible Unicode Cleaner", [
        "Non-breaking spaces", "Zero-width spaces", "Zero-width non-joiners", "Zero-width joiners",
        "Byte order marks", "Soft hyphens", "Non-breaking hyphens",
    ]),
    ("TTF-02_PunctuationNormalizer.docx", "Punctuation Normalizer", [
        "Curly double quotes", "Curly single quotes", "Em dashes", "En dashes", "Ellipses",
    ]),
    ("TTF-03_SpacingFixer.docx", "Spacing Fixer", [
        "Multiple spaces", "Leading and trailing spaces", "Spaces before punctuation",
        "Spacing after punctuation", "Extra blank lines",
    ]),
    ("TTF-04_CapitalizationFixer.docx", "Capitalization Fixer", [
        "Sentence boundaries", "Abbreviations", "Acronyms", "Names and brands", "Heading context",
    ]),
    ("TTF-05_ListNormalizer.docx", "List Normalizer", [
        "Manual hyphen bullets", "Manual asterisk bullets", "Manual numbered items", "List indentation",
    ]),
    ("TTF-06_ParagraphStructureFixer.docx", "Paragraph Structure Fixer", [
        "Consecutive empty paragraphs", "Soft returns to paragraph marks", "Paragraph spacing",
        "Manual tab indentation", "Manual space indentation",
    ]),
    ("TTF-07_SoftReturnConverter.docx", "Soft Return Converter", [
        "Soft returns to paragraph marks", "Paragraph marks to soft returns",
    ]),
    ("TTF-08_TrimEndOfDoc.docx", "Trim End of Doc", ["Trailing empty paragraphs"]),
    ("TTF-09_TableCleaner.docx", "Table Cleaner", [
        "Empty table rows", "Empty table columns", "Cell padding", "Direct cell formatting",
        "Border normalization", "Border removal", "Single-column table to text",
    ]),
    ("TTF-10_HeaderFooterStandardizer.docx", "Header / Footer Standardizer", [
        "Header formatting", "Footer formatting", "Header and footer font", "Header and footer spacing",
        "Header and footer alignment", "Unlink sections", "Clear header and footer content",
    ]),
    ("TTF-11_FootnoteEndnoteRemover.docx", "Footnote / Endnote Remover", [
        "Remove footnotes", "Remove endnotes", "Keep note text inline",
    ]),
    ("TTF-12_ObjectRemover.docx", "Object Remover", [
        "Pictures", "Text boxes", "Frames", "Horizontal lines", "Form controls", "Hidden text", "Tables",
    ]),
    ("TTF-13_FontNormalizer.docx", "Font Normalizer", [
        "Font face", "Font size", "Direct bold", "Direct italic", "Font color",
    ]),
    ("TTF-14_FormattingCleaner.docx", "Formatting Cleaner", [
        "Direct character formatting", "Direct paragraph formatting", "Preserve highlighting",
        "Quick emphasis", "Thorough emphasis", "Strip emphasis",
    ]),
    ("TTF-15_DuplicateParagraphDetector.docx", "Duplicate Paragraph Detector", [
        "Exact duplicate removal", "Normalized duplicate removal", "Fuzzy duplicate removal",
    ]),
    ("TTF-16_MetaDataSuite.docx", "MetaDataSuite", [
        "Document properties", "Personal information", "Sharing properties", "Author identity",
    ]),
    ("TTF-17_FinalReview.docx", "Final Review", [
        "Remove comments", "Accept tracked changes", "Minimal author cleanup",
        "Normal author cleanup", "Broad author cleanup",
    ]),
    ("TTF-18_HyperlinkCleaner.docx", "Hyperlink Cleaner", [
        "Hide hyperlink styling", "Remove hyperlinks",
    ]),
    ("TTF-19_StyleCleanup.docx", "Style Cleanup", [
        "Unused custom styles", "Style variant remapping",
    ]),
    ("TTF-20_BreakNormalizer.docx", "Break Normalizer", [
        "Consecutive section breaks", "Consecutive page breaks", "Next Page conversion",
        "Continuous conversion", "Even Page conversion", "Odd Page conversion",
    ]),
]


def package_names(path: Path) -> set[str]:
    with ZipFile(path) as package:
        return set(package.namelist())


def package_part(path: Path, name: str) -> str:
    with ZipFile(path) as package:
        return package.read(name).decode("utf-8", errors="ignore")


def doc_text(path: Path) -> str:
    with ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        header_footer_parts = [
            name for name in package.namelist()
            if name.startswith("word/header") or name.startswith("word/footer")
        ]
        roots = [root] + [ET.fromstring(package.read(name)) for name in header_footer_parts]

    lines = []
    for part_root in roots:
        for paragraph in part_root.findall(".//w:p", NS):
            text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))
            if text:
                lines.append(text)
    return "\n".join(lines)


def file_hash(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


class TestingFilesAlignmentTests(unittest.TestCase):
    def test_fixture_names_and_titles_follow_live_launcher_order(self):
        expected_names = [name for name, _, _ in EXPECTED_FIXTURES]
        self.assertEqual(expected_names, sorted(path.name for path in ROOT_TESTS.glob("TTF-*.docx")))
        self.assertEqual(expected_names, sorted(path.name for path in PRACTICE_TESTS.glob("TTF-*.docx")))

        for index, (filename, tool_name, _) in enumerate(EXPECTED_FIXTURES, start=1):
            for folder in (ROOT_TESTS, PRACTICE_TESTS):
                with self.subTest(folder=folder, filename=filename):
                    title_text = f"TTF-{index:02d} | {tool_name}"
                    title_sources = doc_text(folder / filename) + package_part(folder / filename, "word/document.xml")
                    self.assertIn(title_text, title_sources)

    def test_old_fixture_files_and_backups_do_not_survive(self):
        allowed_parents = {ROOT_TESTS.resolve(), PRACTICE_TESTS.resolve()}
        all_ttf_files = list(ROOT.rglob("TTF-*.docx"))
        self.assertEqual(40, len(all_ttf_files))
        for path in all_ttf_files:
            with self.subTest(path=path):
                self.assertIn(path.parent.resolve(), allowed_parents)
        self.assertEqual([], list(ROOT_TESTS.glob("~$*.docx")))
        self.assertEqual([], list(PRACTICE_TESTS.glob("~$*.docx")))

    def test_documents_have_only_title_bitmap_legend_and_color_pairs(self):
        prohibited = [
            "PASS / FAIL", "CHECKLIST", "HOW TO USE", "BEFORE", "EXPECTED AFTER",
            "TEST:", "Info:", "Result:", "Tool Testing File",
        ]
        legend_phrases = ["Tan: This is how", "Green: The tan sample"]

        for filename, _, markers in EXPECTED_FIXTURES:
            path = ROOT_TESTS / filename
            text = doc_text(path)
            xml = package_part(path, "word/document.xml")
            names = package_names(path)
            with self.subTest(filename=filename):
                for phrase in prohibited + legend_phrases:
                    self.assertNotIn(phrase, text)
                for marker in markers:
                    self.assertGreaterEqual(text.count(marker), 2, marker)
                self.assertGreaterEqual(xml.count('w:fill="FFF2CC"'), len(markers))
                self.assertGreaterEqual(xml.count('w:fill="E2F0D9"'), len(markers))
                self.assertTrue(any(name.startswith("word/media/") for name in names))

    def test_root_and_practice_fixtures_are_identical(self):
        for filename, _, _ in EXPECTED_FIXTURES:
            with self.subTest(filename=filename):
                self.assertEqual(file_hash(ROOT_TESTS / filename), file_hash(PRACTICE_TESTS / filename))

    def test_structural_fixtures_use_real_word_structures(self):
        list_xml = package_part(ROOT_TESTS / "TTF-05_ListNormalizer.docx", "word/document.xml")
        self.assertIn('w:pStyle w:val="ListBullet"', list_xml)
        self.assertIn('w:pStyle w:val="ListNumber"', list_xml)

        table_xml = package_part(ROOT_TESTS / "TTF-09_TableCleaner.docx", "word/document.xml")
        self.assertGreaterEqual(table_xml.count("<w:tbl>"), 2)

        header_footer_names = package_names(ROOT_TESTS / "TTF-10_HeaderFooterStandardizer.docx")
        self.assertTrue(any(name.startswith("word/header") for name in header_footer_names))
        self.assertTrue(any(name.startswith("word/footer") for name in header_footer_names))

        note_names = package_names(ROOT_TESTS / "TTF-11_FootnoteEndnoteRemover.docx")
        self.assertIn("word/footnotes.xml", note_names)
        self.assertIn("word/endnotes.xml", note_names)
        self.assertIn("<w:footnoteRef", package_part(ROOT_TESTS / "TTF-11_FootnoteEndnoteRemover.docx", "word/footnotes.xml"))
        self.assertIn("<w:endnoteRef", package_part(ROOT_TESTS / "TTF-11_FootnoteEndnoteRemover.docx", "word/endnotes.xml"))

        object_path = ROOT_TESTS / "TTF-12_ObjectRemover.docx"
        object_xml = package_part(object_path, "word/document.xml")
        self.assertIn("<w:vanish", object_xml)
        self.assertIn("CleanupSuiteFixtureFormControl", object_xml)
        self.assertIn("<w:tbl>", object_xml)
        self.assertTrue("<w:drawing" in object_xml or "<w:pict" in object_xml)
        self.assertIn("textbox", object_xml)
        self.assertIn("o:hr", object_xml)

        review_path = ROOT_TESTS / "TTF-17_FinalReview.docx"
        review_names = package_names(review_path)
        review_xml = package_part(review_path, "word/document.xml")
        self.assertIn("word/comments.xml", review_names)
        self.assertIn("<w:ins", review_xml)

        hyperlink_xml = package_part(ROOT_TESTS / "TTF-18_HyperlinkCleaner.docx", "word/document.xml")
        self.assertIn("<w:hyperlink", hyperlink_xml)

        break_xml = package_part(ROOT_TESTS / "TTF-20_BreakNormalizer.docx", "word/document.xml")
        self.assertIn('w:type="page"', break_xml)
        self.assertGreaterEqual(break_xml.count("<w:sectPr"), 2)


if __name__ == "__main__":
    unittest.main()
