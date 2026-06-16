# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Christopher Travis Massey
# See LICENSE for full terms.

"""
assemble.py  --  Rebuild the distributable VBA_Cleanup_tool.txt from the
                 per-component source files in src/, then run the build-gate
                 validators.  Exits 1 on any validation failure.

Usage:
    python assemble.py [--src=<dir>] [--out=<file>] [--no-validate]

    --src          source tree root  (default: src/)
    --out          output .txt path  (default: VBA_Cleanup_tool.txt)
    --no-validate  skip build-gate validators (not recommended)

How it works
------------
Reads src/manifest.txt line by line.  Each directive produces a CRLF chunk
that is concatenated in order into the final output:

    VERBATIM  path
        Read the file, convert LF -> CRLF, append.

    BUILDER  path  func_name
        Read the file (plain VBA, LF endings), run vba_to_builder to produce
        the Private Function func_name() As String block, append.

    XML_BUILDER  path  func_name
        Same as BUILDER but source is plain XML (ribbon special case).

    SEP  <base64>
        Decode the base64 blob (a raw CRLF block), append verbatim.
        These carry the comment lines that live between builders.

After assembly the full build-gate is run via vbaeval.run_all().  Any failure
prints the report and exits 1 so the error is visible in CI / pre-commit hooks.

The output file is only written if validation passes (or --no-validate is set).
"""

import sys
import os
import base64
import shutil

DEFAULT_SRC  = "src"
DEFAULT_OUT  = "VBA_Cleanup_tool.txt"
PRACTICE_DIR = "Practice - Try CleanupSuite Here"

# ---------------------------------------------------------------- vba_to_builder
# (inline copy so assemble.py has no runtime dependency on vbaeval.py being
#  importable — it may live in a different directory)
def vba_to_builder(vba, func):
    """Plain VBA (LF endings) -> a CRLF `Private Function func() As String` block."""
    lines = vba.replace("\r\n", "\n").split("\n")
    o = [f"Private Function {func}() As String", "    Dim s As String"]
    for k, ln in enumerate(lines):
        suf = '"' if k == len(lines) - 1 else '" & vbCrLf'
        o.append('    s = s & "' + ln.replace('"', '""') + suf)
    o.append(f"    {func} = s")
    o.append("End Function")
    return "\r\n".join(o)

# ---------------------------------------------------------------- helpers
def read_lf(path):
    """Read a source file, normalise to LF endings."""
    with open(path, "r", newline="", encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")

def read_verbatim_as_crlf(path):
    """Read a source file, convert LF -> CRLF for the distributable."""
    text = read_lf(path)
    return text.replace("\n", "\r\n")

def parse_manifest(manifest_path):
    """Yield (type, args) tuples for each non-comment, non-blank line."""
    with open(manifest_path, "r", newline="\n", encoding="utf-8") as f:
        for raw in f:
            line = raw.rstrip("\n")
            if not line or line.startswith("#"):
                continue
            parts = line.split(None, 2)   # max 3 parts: TYPE path [func]
            yield parts[0], parts[1:]


def sync_practice_copy(out_file, repo_root=None):
    """Refresh the practice distributable when the root bundle is rebuilt."""
    repo_root = repo_root or os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(out_file) != DEFAULT_OUT:
        return None
    practice_dir = os.path.join(repo_root, PRACTICE_DIR)
    if not os.path.isdir(practice_dir):
        return None
    practice_out = os.path.join(practice_dir, DEFAULT_OUT)
    if os.path.abspath(out_file) == os.path.abspath(practice_out):
        return None
    shutil.copyfile(out_file, practice_out)
    return practice_out

# ---------------------------------------------------------------- assemble
def assemble(src_dir, out_file, validate=True, repo_root=None):
    manifest_path = os.path.join(src_dir, "manifest.txt")
    if not os.path.exists(manifest_path):
        sys.exit(f"ERROR: manifest not found: {manifest_path}")

    chunks = []

    for directive, args in parse_manifest(manifest_path):
        if directive == "VERBATIM":
            path = os.path.join(src_dir, args[0])
            if not os.path.exists(path):
                sys.exit(f"ERROR: VERBATIM file not found: {path}")
            chunk = read_verbatim_as_crlf(path)
            chunks.append(chunk)

        elif directive in ("BUILDER", "XML_BUILDER"):
            if len(args) < 2:
                sys.exit(f"ERROR: {directive} requires path and func_name")
            path      = os.path.join(src_dir, args[0])
            func_name = args[1]
            if not os.path.exists(path):
                sys.exit(f"ERROR: {directive} file not found: {path}")
            source = read_lf(path)
            chunk  = vba_to_builder(source, func_name)
            chunks.append(chunk)

        elif directive == "SEP":
            # args[0] may be absent (blank-only separator encodes to empty string
            # and the manifest line is just "SEP  " with no trailing token).
            b64 = args[0] if args else ""
            try:
                sep_bytes = base64.b64decode(b64) if b64 else b""
                chunk = sep_bytes.decode("utf-8")
            except Exception as e:
                sys.exit(f"ERROR: SEP base64 decode failed: {e}")
            chunks.append(chunk)

        else:
            sys.exit(f"ERROR: unknown manifest directive: {directive!r}")

    # Join: each chunk already ends without a trailing CRLF (VERBATIM files end
    # at their last line; builders end at End Function; SEPs end at their last
    # comment line).  We join with CRLF between chunks.
    content = "\r\n".join(chunks)

    # Ensure file ends with a single trailing CRLF (matches original)
    content = content.rstrip("\r\n") + "\r\n"

    # ---- validate before writing ------------------------------------
    if validate:
        # Import vbaeval from the same directory as this script, or from cwd.
        script_dir = os.path.dirname(os.path.abspath(__file__))
        for search in [script_dir, os.getcwd()]:
            candidate = os.path.join(search, "vbaeval.py")
            if os.path.exists(candidate):
                _ns = {}
                with open(candidate, "r", encoding="utf-8") as f:
                    exec(compile(f.read(), candidate, "exec"), _ns)
                run_all = _ns["run_all"]
                break
        else:
            print("WARNING: vbaeval.py not found — skipping validation")
            run_all = None

        if run_all:
            print("Running build-gate validators...")
            ok, report = run_all(content)
            for line in report:
                print(f"  {line}")
            if not ok:
                print("\nBUILD FAILED — distributable not written.")
                sys.exit(1)

    # ---- write output -----------------------------------------------
    with open(out_file, "w", newline="", encoding="utf-8") as f:
        f.write(content)

    lines_n = content.count("\r\n")
    print(f"Assembled -> {out_file}  ({lines_n} lines, {len(content)} bytes)")
    practice_out = sync_practice_copy(out_file, repo_root=repo_root)
    if practice_out:
        print(f"Synced practice bundle -> {practice_out}")
    return content

# ---------------------------------------------------------------- CLI
if __name__ == "__main__":
    argv = sys.argv[1:]
    src_dir   = next((a.split("=",1)[1] for a in argv if a.startswith("--src=")),  DEFAULT_SRC)
    out_file  = next((a.split("=",1)[1] for a in argv if a.startswith("--out=")),  DEFAULT_OUT)
    no_validate = "--no-validate" in argv

    if not os.path.isdir(src_dir):
        sys.exit(f"ERROR: src directory not found: {src_dir}")

    print(f"Assembling from {src_dir}/ -> {out_file}")
    assemble(src_dir, out_file, validate=not no_validate)
