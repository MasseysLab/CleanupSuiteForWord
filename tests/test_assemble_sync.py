from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SPEC = spec_from_file_location("cleanupsuite_assemble", ROOT / "assemble.py")
ASSEMBLE = module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(ASSEMBLE)


def read_exact(path: Path) -> str:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return handle.read()


class AssembleSyncTests(unittest.TestCase):
    def test_manifest_parser_ignores_crlf_whitespace_only_lines(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest = Path(temp_dir) / "manifest.txt"
            manifest.write_bytes(
                b"# Windows checkout\r\n\r\n  \r\nVERBATIM sample.bas\r\nSEP\r\n"
            )

            self.assertEqual(
                list(ASSEMBLE.parse_manifest(str(manifest))),
                [("VERBATIM", ["sample.bas"]), ("SEP", [])],
            )

    def test_rebuilding_root_bundle_refreshes_practice_copy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            src_dir = temp_root / "src"
            practice_dir = temp_root / "Practice - Try CleanupSuite Here"
            src_dir.mkdir()
            practice_dir.mkdir()

            (src_dir / "manifest.txt").write_text(
                "VERBATIM sample.bas\n",
                encoding="utf-8",
                newline="\n",
            )
            (src_dir / "sample.bas").write_text(
                "Option Explicit\nPublic Sub Hello()\nEnd Sub\n",
                encoding="utf-8",
                newline="\n",
            )

            out_file = temp_root / "VBA_Cleanup_tool.txt"
            ASSEMBLE.assemble(str(src_dir), str(out_file), validate=False, repo_root=str(temp_root))

            practice_out = practice_dir / "VBA_Cleanup_tool.txt"
            self.assertTrue(practice_out.exists())
            self.assertEqual(read_exact(out_file), read_exact(practice_out))
