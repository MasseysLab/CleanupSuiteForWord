# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Christopher Travis Massey
# See LICENSE for full terms.

"""
explode.py  --  Split the distributable VBA_Cleanup_tool.txt into per-component
                source files and emit src/manifest.txt.

Usage:
    python explode.py [--src=<dir>] [--file=<txt>]

    --file  path to the distributable .txt  (default: VBA_Cleanup_tool.txt)
    --src   destination source tree root    (default: src/)

What it produces
----------------
src/
  manifest.txt                   assembly-order file consumed by assemble.py
  installer/installer.bas        outer host module verbatim (lines 1..end-of-ControlTypeByName)
  installer/checklist.bas        ManualChecklistText emitted VBA (plain, no builder strings)
  modules/modCleanupHelpers.bas  emitted VBA for modCleanupHelpers
  modules/modCleanupLauncher.bas emitted VBA for modCleanupLauncher
  forms/frmXxx.bas               emitted VBA for each UserForm (one file per form)
  ribbon/ribbon.xml              plain ribbon XML (extracted from RIBBON_XML builder)
  ribbon/inject.bas              InjectRibbonXml + Wait verbatim

All .bas/.xml source files use LF line endings (the natural authoring format).
assemble.py converts back to CRLF when building the distributable.

The manifest uses a simple line format:
    TYPE  path  [builder_func]
where TYPE is one of:
    VERBATIM   -- paste file content as-is into distributable (CRLF-converted)
    BUILDER    -- convert file -> builder function via vba_to_builder
    XML_BUILDER-- like BUILDER but source is plain XML (ribbon special case)
    SEP        -- literal separator lines stored inline (base64-encoded CRLF block)

SEP lines carry the verbatim CRLF text between the installer body and the first
builder, and between consecutive builders, encoded as base64 so the manifest
stays a plain line-oriented file.
"""

import re
import sys
import os
import base64

DEFAULT_FILE = "VBA_Cleanup_tool.txt"
DEFAULT_SRC  = "src"

# ---------------------------------------------------------------- helpers
def load(path):
    with open(path, "r", newline="", encoding="utf-8") as f:
        return f.read()

def write_lf(path, text):
    """Write text with LF endings (normalise from whatever the source uses)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="\n", encoding="utf-8") as f:
        f.write(text.replace("\r\n", "\n"))

def encode_sep(crlf_text):
    """Base64-encode a CRLF separator block for embedding in manifest."""
    return base64.b64encode(crlf_text.encode("utf-8")).decode("ascii")

# ---------------------------------------------------------------- extract_code (accumulator-aware)
def eval_concat(expr):
    n = len(expr); buf = ""; inq = False; i = 0; tokens = []
    while i < n:
        ch = expr[i]
        if ch == '"':
            inq = not inq; buf += ch; i += 1
        elif ch == "&" and not inq:
            tokens.append(buf); buf = ""; i += 1
        else:
            buf += ch; i += 1
    tokens.append(buf)
    out = ""
    for t in tokens:
        t = t.strip()
        if not t:
            continue
        if t.startswith('"') and t.endswith('"'):
            out += t[1:-1].replace('""', '"')
        elif t in ("vbCrLf", "vbNewLine", "vbCr", "vbLf"):
            out += "\n"
        elif t == "vbTab":
            out += "\t"
    return out

def extract_code(name, c):
    start = c.find(f"Private Function {name}() As String")
    if start < 0:
        return None
    end = c.find("\r\nEnd Function", start)
    body = c[start:end]
    acc = "s"
    for ln in body.split("\r\n"):
        st = ln.strip()
        m = re.match(r"(\w+)\s*=\s*(\w+)\s*$", st)
        if m and m.group(1) == name:
            acc = m.group(2)
            break
    prefix = acc + " = " + acc + " & "
    out = ""
    for ln in body.split("\r\n"):
        s = ln.strip()
        if s.startswith(prefix):
            out += eval_concat(s[len(prefix):])
    return out

# ---------------------------------------------------------------- main
def explode(src_file, src_dir):
    c = load(src_file)
    lines = c.split("\r\n")

    # ---- locate builder spans ----------------------------------------
    builder_spans = []   # (name, start_line_1based, end_line_1based)
    in_builder = False
    b_start = None
    b_name  = None
    for i, ln in enumerate(lines, 1):
        s = ln.strip()
        m = re.match(r"Private Function (\w+)\(\) As String", s)
        if m:
            in_builder = True; b_start = i; b_name = m.group(1)
        elif s == "End Function" and in_builder:
            builder_spans.append((b_name, b_start, i))
            in_builder = False

    # ---- locate outer module end (end of ControlTypeByName) ----------
    outer_end = None
    in_fn = False
    for i, ln in enumerate(lines, 1):
        s = ln.strip()
        if s == "Private Function ControlTypeByName(ctrlName As String) As String":
            in_fn = True
        elif s == "End Function" and in_fn:
            outer_end = i; in_fn = False; break
    if outer_end is None:
        sys.exit("ERROR: could not locate end of ControlTypeByName in installer")

    # ---- verify InjectRibbonXml sub exists (structural sanity check) ---
    if not any(ln.strip() == "Public Sub InjectRibbonXml(filePath As String)"
               for ln in lines):
        sys.exit("ERROR: could not locate InjectRibbonXml in source")

    # ---- locate RIBBON_XML span (needed to slice inject_lines) ----------
    ribbon_xml_span = next(sp for sp in builder_spans if sp[0] == "RIBBON_XML")
    inject_lines = lines[ribbon_xml_span[2]:]
    while inject_lines and not inject_lines[-1].strip():
        inject_lines.pop()

    # ---- write installer.bas (lines 1..outer_end) --------------------
    installer_text = "\r\n".join(lines[:outer_end])
    write_lf(os.path.join(src_dir, "installer", "installer.bas"), installer_text)
    print(f"  wrote installer/installer.bas  ({outer_end} lines)")

    # ---- write inject.bas -----------------------------------------------
    # inject.bas = everything after the RIBBON_XML End Function to end of file,
    # stripping any trailing blank lines.  The leading blank line inside the file
    # provides the separation when joined to the RIBBON_XML builder chunk.

    write_lf(os.path.join(src_dir, "ribbon", "inject.bas"),
             "\r\n".join(inject_lines))
    print(f"  wrote ribbon/inject.bas  ({len(inject_lines)} lines)")

    # ---- extract and write each builder's source ---------------------
    # Determine destination path and sub-dir for each builder
    def dest_for(name):
        if name == "ManualChecklistText":
            return "installer/checklist.bas", "installer"
        if name == "RIBBON_XML":
            return "ribbon/ribbon.xml", "ribbon"
        if name.startswith("mod"):
            stem = name[:-5] if name.endswith("_Code") else name
            return f"modules/{stem}.bas", "modules"
        if name.startswith("frm"):
            stem = name[:-5] if name.endswith("_Code") else name
            return f"forms/{stem}.bas", "forms"
        return f"misc/{name}.bas", "misc"

    written = {}  # name -> rel_path
    for name, s, e in builder_spans:
        if name == "RIBBON_XML":
            # extract as plain XML
            xml_text = extract_code(name, c)
            rel_path, _ = dest_for(name)
            write_lf(os.path.join(src_dir, rel_path), xml_text)
            print(f"  wrote {rel_path}  ({len(xml_text)} chars)")
        else:
            code = extract_code(name, c)
            rel_path, _ = dest_for(name)
            write_lf(os.path.join(src_dir, rel_path), code)
            lines_n = code.count("\n") + 1
            print(f"  wrote {rel_path}  ({lines_n} lines)")
        written[name] = dest_for(name)[0]

    # ---- build manifest.txt ------------------------------------------
    manifest_lines = [
        "# manifest.txt -- assembly order for assemble.py",
        "# Generated by explode.py -- do not edit the SEP entries by hand.",
        "#",
        "# FORMAT:  TYPE  path  [builder_func]",
        "#   VERBATIM    paste file verbatim (CRLF-converted) into distributable",
        "#   BUILDER     convert source file -> builder function (uses vba_to_builder)",
        "#   XML_BUILDER like BUILDER but source is plain XML",
        "#   SEP         literal CRLF separator block (base64-encoded, inline)",
        "#",
        "# SEP lines: SEP <base64>",
        "",
        "VERBATIM  installer/installer.bas",
        "",
    ]

    # Now emit SEP + BUILDER for each builder in order
    prev_end_1based = outer_end
    for name, s, e in builder_spans:
        # separator block = lines between prev_end and this builder start
        sep_lines = lines[prev_end_1based:s-1]
        sep_crlf  = "\r\n".join(sep_lines)
        # Always emit SEP — even a blank-only separator (empty string) must be
        # recorded so assemble knows to emit the separating blank line.
        manifest_lines.append(f"SEP  {encode_sep(sep_crlf)}")

        rel_path = written[name]
        if name == "RIBBON_XML":
            manifest_lines.append(f"XML_BUILDER  {rel_path}  {name}")
        else:
            manifest_lines.append(f"BUILDER  {rel_path}  {name}")

        prev_end_1based = e

    # Post-RIBBON_XML tail: the SEP emitted by the loop for the block after
    # frmCleanupSuiteLauncher_Code already carries the ribbon module header
    # comment block verbatim (it's base64-encoded in the manifest).
    # inject.bas itself is the final VERBATIM entry.
    manifest_lines.append("")
    manifest_lines.append("VERBATIM  ribbon/inject.bas")
    manifest_lines.append("")

    manifest_path = os.path.join(src_dir, "manifest.txt")
    os.makedirs(src_dir, exist_ok=True)
    with open(manifest_path, "w", newline="\n", encoding="utf-8") as f:
        f.write("\n".join(manifest_lines) + "\n")
    print(f"  wrote manifest.txt")

    print(f"\nExplode complete -> {src_dir}/")
    print(f"  {len(builder_spans)} builders extracted")
    print(f"  manifest has {len([l for l in manifest_lines if l.startswith(('VERBATIM','BUILDER','XML_BUILDER','SEP'))])} directives")


# ---------------------------------------------------------------- CLI
if __name__ == "__main__":
    argv = sys.argv[1:]
    src_file = next((a.split("=",1)[1] for a in argv if a.startswith("--file=")), DEFAULT_FILE)
    src_dir  = next((a.split("=",1)[1] for a in argv if a.startswith("--src=")),  DEFAULT_SRC)

    if not os.path.exists(src_file):
        sys.exit(f"ERROR: source file not found: {src_file}")

    print(f"Exploding {src_file} -> {src_dir}/")
    explode(src_file, src_dir)
