import importlib.util
import os
import sys
import types
import unittest
from pathlib import Path


DOCUMENTS_SKILL_ROOT = (
    Path.home() / ".codex" / "plugins" / "cache" / "openai-primary-runtime" / "documents"
)


def resolve_render_docx_path() -> Path:
    candidates = sorted(DOCUMENTS_SKILL_ROOT.glob(r"*\skills\documents\render_docx.py"))
    if not candidates:
        raise unittest.SkipTest(
            f"bundled render_docx.py is not available under {DOCUMENTS_SKILL_ROOT}"
        )
    return candidates[-1]


def load_render_docx_module():
    original_pdf2image = sys.modules.get("pdf2image")
    if original_pdf2image is None:
        stub = types.ModuleType("pdf2image")
        stub.convert_from_path = lambda *args, **kwargs: []
        stub.pdfinfo_from_path = lambda *args, **kwargs: {}
        sys.modules["pdf2image"] = stub
    spec = importlib.util.spec_from_file_location("render_docx_module", resolve_render_docx_path())
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    try:
        spec.loader.exec_module(module)
    finally:
        if original_pdf2image is None:
            sys.modules.pop("pdf2image", None)
        else:
            sys.modules["pdf2image"] = original_pdf2image
    return module


def require_helper(module, helper_name: str):
    helper = getattr(module, helper_name, None)
    if helper is None:
        raise unittest.SkipTest(f"{helper_name} is not exposed by this bundled render_docx.py")
    return helper


class RenderDocxWindowsLibreOfficePathTests(unittest.TestCase):
    def test_windows_prefers_full_soffice_console_path(self):
        module = load_render_docx_module()
        original_platform = module.sys.platform
        original_program_files = os.environ.get("ProgramFiles")
        original_program_files_x86 = os.environ.get("ProgramFiles(x86)")
        try:
            module.sys.platform = "win32"
            os.environ["ProgramFiles"] = r"C:\Program Files"
            os.environ["ProgramFiles(x86)"] = r"C:\Program Files (x86)"

            resolve_soffice_command = require_helper(module, "_resolve_soffice_command")
            soffice = resolve_soffice_command()

            self.assertEqual(
                r"C:\Program Files\LibreOffice\program\soffice.com",
                soffice,
            )
        finally:
            module.sys.platform = original_platform
            if original_program_files is None:
                os.environ.pop("ProgramFiles", None)
            else:
                os.environ["ProgramFiles"] = original_program_files
            if original_program_files_x86 is None:
                os.environ.pop("ProgramFiles(x86)", None)
            else:
                os.environ["ProgramFiles(x86)"] = original_program_files_x86

    def test_windows_user_installation_uri_uses_file_triple_slash(self):
        module = load_render_docx_module()
        original_platform = module.sys.platform
        try:
            module.sys.platform = "win32"
            user_installation_uri = require_helper(module, "_user_installation_uri")
            uri = user_installation_uri(
                r"C:\Users\Chris\AppData\Local\Temp\soffice_profile_test"
            )
            self.assertEqual(
                "file:///C:/Users/Chris/AppData/Local/Temp/soffice_profile_test",
                uri,
            )
        finally:
            module.sys.platform = original_platform

    def test_windows_resolves_bundled_poppler_directory(self):
        module = load_render_docx_module()
        original_platform = module.sys.platform
        original_executable = module.sys.executable
        try:
            module.sys.platform = "win32"
            module.sys.executable = (
                r"C:\Users\Chris\.cache\codex-runtimes\codex-primary-runtime"
                r"\dependencies\python\python.exe"
            )
            resolve_poppler_path = require_helper(module, "_resolve_poppler_path")
            poppler_dir = resolve_poppler_path()
            self.assertEqual(
                (
                    r"C:\Users\Chris\.cache\codex-runtimes\codex-primary-runtime"
                    r"\dependencies\native\poppler\Library\bin"
                ),
                poppler_dir,
            )
        finally:
            module.sys.platform = original_platform
            module.sys.executable = original_executable


if __name__ == "__main__":
    unittest.main()
