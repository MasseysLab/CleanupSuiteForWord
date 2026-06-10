from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class LauncherRedesignTests(unittest.TestCase):
    def test_launcher_has_category_labels_and_tool_descriptions(self):
        installer = read("src/installer/installer.bas")

        for category in [
            "Text and Characters",
            "Paragraphs, Breaks, and Lists",
            "Layout and Document Structure",
            "Formatting, Links, and Styles",
            "Review, Privacy, and Removals",
        ]:
            self.assertIn(category, installer)

        for label in [
            "lblDescUnicode",
            "lblDescDuplicate",
            "lblDescTableClean",
            "lblDescMetadata",
            "lblDescObjectRemover",
        ]:
            self.assertIn(label, installer)

    def test_launcher_layout_is_row_based_instead_of_flat_button_grid(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("LayoutLauncherCategory", installer)
        self.assertIn("LayoutLauncherToolRow", installer)
        self.assertNotIn('Array("cmdUnicode", "cmdHelpUnicode")', installer)

    def test_launcher_milestone_version_is_060(self):
        launcher = read("src/modules/modCleanupLauncher.bas")
        versioning = read("VERSIONING.md")

        self.assertIn('Public Const SUITE_VERSION As String = "0.6.0"', launcher)
        self.assertIn("## Current Version\n\n`0.6.0`", versioning)


if __name__ == "__main__":
    unittest.main()
