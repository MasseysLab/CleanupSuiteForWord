from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def procedure_blocks(source: str):
    pattern = re.compile(
        r"(?im)^\s*(?:Private|Public)?\s*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)[^\r\n]*"
    )
    matches = list(pattern.finditer(source))
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(source)
        yield match.group(1), source[match.start():end]


def dim_names(line: str):
    stripped = line.split("'", 1)[0].strip()
    if not stripped.lower().startswith("dim "):
        return []
    declaration = stripped[4:]
    candidate = declaration.split(",", 1)[0].strip()
    names = [re.split(r"\s+As\s+|\s*:", candidate, maxsplit=1, flags=re.IGNORECASE)[0].strip()]
    return [name.lower() for name in names if name]


class ToolCompileRegressionTests(unittest.TestCase):
    def test_spacing_tool_has_no_duplicate_local_dim_names(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        for proc_name, block in procedure_blocks(source):
            seen = set()
            duplicates = []
            for line in block.splitlines():
                for name in dim_names(line):
                    if name in seen:
                        duplicates.append(name)
                    seen.add(name)
            self.assertEqual([], duplicates, f"{proc_name} has duplicate local declarations")

    def test_table_cleaner_uses_valid_word_range_formatting_calls(self):
        source = read("src/forms/frmTableCleaner.bas")

        self.assertNotIn("tbl.Range.ClearFormatting", source)
        self.assertIn("tbl.Range.Font.Reset", source)
        self.assertIn("tbl.Range.ParagraphFormat.Reset", source)

    def test_spacing_tool_does_not_use_invalid_word_wildcard_punctuation_pattern(self):
        source = read("src/forms/frmSpacingCleanup.bas")

        self.assertNotIn('" ([.,;:!?])"', source)
        self.assertNotIn(".Replacement.Text = \"\\1\"", source)
        self.assertIn("beforePunctPatterns = Array(\" .\", \" ,\", \" ;\", \" :\", \" !\", \" ?\")", source)
        self.assertIn("previewBeforePunctPatterns = Array(\" .\", \" ,\", \" ;\", \" :\", \" !\", \" ?\")", source)

    def test_metadata_preview_has_no_blank_line_after_heading(self):
        source = read("src/forms/frmMetadataScrubber.bas")

        self.assertIn('"Current document metadata:" & vbCrLf', source)
        self.assertNotIn('"Current document metadata:" & vbCrLf & vbCrLf', source)
