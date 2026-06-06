# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Christopher Travis Massey
# See LICENSE for full terms.

"""
vbaeval.py  --  authoring/validation helper for the Word VBA Cleanup Suite.

The suite ships as one CRLF .txt that self-injects forms+modules at runtime by
destringifying builder functions.  The canonical form uses `s` as the accumulator:

    Private Function frmXxx_Code() As String
        Dim s As String
        s = s & "Option Explicit" & vbCrLf
        s = s & "..." & vbCrLf
        s = s & "last content line has NO trailing & vbCrLf"
        frmXxx_Code = s
    End Function

Some builders use a differently-named accumulator (e.g. ManualChecklistText uses
`checklistText`).  extract_code and unhandled_concat_tokens detect the accumulator
variable from the return-assignment line (`<funcname> = <var>`) rather than
hard-coding `s`, so all builders are handled correctly regardless of variable name.
vba_to_builder always emits the canonical `s`-accumulator form.

This module lets you:
  * read the .txt (content)
  * extract_code(name)         -> the injected VBA a builder emits (\\n-separated)
  * vba_to_builder(vba, func)  -> the reverse: real VBA -> builder function text
  * replace_builder(c,func,vba)-> swap a builder inside a content string
and run the build-gate validators:
  * balance(code)              -> statement nesting must zero out, never go negative
  * single_line_if_traps(code) -> single-line `If ... Then <block>` = compile death
  * control_wiring(content)    -> every control used in a form is declared in its
                                  ControlsForForm Array(...)
  * unterminated_strings(content)
  * round_trip_ok(content, name)

CRLF is preserved on read/write by using newline=''.  extract_code returns text
with '\\n' separators (vbCrLf -> \\n); feed \\n-normalised VBA back to vba_to_builder.

CLI:
    python vbaeval.py validate            # run all gates, exit 1 on any failure
    python vbaeval.py list                # list builder names
    python vbaeval.py show frmXxx_Code    # print the injected VBA for one builder
Set the source file with --file (default: VBA_Cleanup_tool.txt in cwd).
"""

import re
import sys

DEFAULT_FILE = "VBA_Cleanup_tool.txt"

# ---------------------------------------------------------------- load
def load(path=DEFAULT_FILE):
    with open(path, "r", newline="", encoding="utf-8") as f:
        return f.read()

def save(content, path=DEFAULT_FILE):
    # newline='' keeps the CRLFs already embedded in `content` intact.
    with open(path, "w", newline="", encoding="utf-8") as f:
        f.write(content)

# Loaded lazily so importing the module never fails if the file is absent.
try:
    content = load()
except OSError:
    content = ""

# ---------------------------------------------------------------- string tools
def eval_concat(expr):
    """Evaluate the RHS of `s = s & <expr>`: a chain of "literals" and vb consts
    joined by &.  Splits on top-level & only (ignores & inside string literals),
    un-doubles "" -> ", maps vbCrLf/vbCr/vbLf/vbNewLine -> \\n, vbTab -> \\t."""
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
        # any other token (a variable, function call) is ignored on purpose;
        # builders are pure literal concatenation, so this never triggers.
    return out

def builder_names(c=None):
    c = content if c is None else c
    return [b for b in re.findall(r"Private Function (\w+)\(\) As String", c)
            if b != "RIBBON_XML"]

def extract_code(name, c=None):
    """Return the injected VBA a builder emits, with '\\n' separators.

    Detects the accumulator variable from the return assignment line
    (`<funcname> = <var>`) rather than hard-coding 's'.  This handles
    builders like ManualChecklistText that use a differently-named
    accumulator (e.g. `checklistText = checklistText & ...`).
    Falls back to 's' if no return assignment is found.
    """
    c = content if c is None else c
    start = c.find(f"Private Function {name}() As String")
    if start < 0:
        return None
    end = c.find("\r\nEnd Function", start)
    body = c[start:end]
    # Detect accumulator: scan for the line `<funcname> = <varname>` with
    # nothing else on the line (the return assignment).
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

def vba_to_builder(vba, func):
    """Real VBA -> a `Private Function func() As String` builder.
    Doubles internal quotes; omits `& vbCrLf` on the LAST content line so the
    round-trip is exact (the suite's builders never trail the final line)."""
    # exact inverse of extract_code: preserve leading/trailing blank lines,
    # so a re-extract reproduces the emitted code byte-for-byte.
    lines = vba.replace("\r\n", "\n").split("\n")
    o = [f"Private Function {func}() As String", "    Dim s As String"]
    for k, ln in enumerate(lines):
        suf = '"' if k == len(lines) - 1 else '" & vbCrLf'
        o.append('    s = s & "' + ln.replace('"', '""') + suf)
    o.append(f"    {func} = s")
    o.append("End Function")
    return "\r\n".join(o)

def replace_builder(c, func, vba):
    """Replace builder `func` inside content string `c` with one built from `vba`."""
    s = c.find(f"Private Function {func}() As String")
    if s < 0:
        raise ValueError(f"builder not found: {func}")
    e = c.find("\r\nEnd Function", s) + len("\r\nEnd Function")
    return c[:s] + vba_to_builder(vba, func) + c[e:]

# ---------------------------------------------------------------- validators
def strip_sc(line):
    """Strip string literals (-> "") and trailing ' comments, so colon-splitting
    and keyword matching never trip on text inside quotes or comments."""
    out = []; i = 0; n = len(line)
    while i < n:
        if line[i] == '"':
            i += 1
            while i < n:
                if line[i] == '"':
                    if i + 1 < n and line[i + 1] == '"':
                        i += 2; continue
                    i += 1; break
                i += 1
            out.append('""')
        elif line[i] == "'":
            break
        else:
            out.append(line[i]); i += 1
    return "".join(out)

def statements(code):
    for raw in code.replace("\r\n", "\n").split("\n"):
        for st in strip_sc(raw).split(":"):
            st = st.strip()
            if st:
                yield st

def balance(code):
    """Return (final_counts, min_counts). Healthy code: all final == 0 and no
    min < 0.  Keys: sub fn with if for do sel."""
    d = {"sub": 0, "fn": 0, "with": 0, "if": 0, "for": 0, "do": 0, "sel": 0}
    mn = dict(d)
    for st in statements(code):
        l = st.lower()
        if re.match(r"(public |private |friend |static )*sub \w+", l): d["sub"] += 1
        elif l == "end sub": d["sub"] -= 1
        if re.match(r"(public |private |friend |static )*function \w+", l): d["fn"] += 1
        elif l == "end function": d["fn"] -= 1
        if re.match(r"with\b", l): d["with"] += 1
        elif l == "end with": d["with"] -= 1
        if re.match(r"if\b.*\bthen$", l): d["if"] += 1     # block-If only (ends in Then)
        elif l == "end if": d["if"] -= 1
        if re.match(r"for\b", l): d["for"] += 1
        elif l == "next" or l.startswith("next "): d["for"] -= 1
        if re.match(r"do\b", l): d["do"] += 1
        elif l.startswith("loop"): d["do"] -= 1
        if re.match(r"select case\b", l): d["sel"] += 1
        elif l == "end select": d["sel"] -= 1
        for k in d:
            if d[k] < mn[k]:
                mn[k] = d[k]
    return d, mn

_BADBLOCK = re.compile(r"\b(End If|End With|End Select|Next\b|Loop\b|Select Case)\b", re.I)
_OPENBLOCK = re.compile(r"(?<![A-Za-z])(With |For |Do )", re.I)
_EXIT = re.compile(r"\bExit (Do|For|Sub|Function)\b", re.I)

def single_line_if_traps(code):
    """Return a list of offending lines: a single-line `If ... Then <X>` where X
    contains a block construct.  Only simple statements may follow Then on one
    line; a block (With/For/Do or a stray End*/Next/Loop) fails to compile.
    Exit Do/For/Sub/Function are allowed and stripped before testing."""
    bad = []
    for raw in code.split("\n"):
        c = strip_sc(raw).strip()
        m = re.match(r"(.*?:)?\s*if\b.*?\bthen\b(.+)$", c, re.I)
        if m and m.group(2).strip():
            tail = _EXIT.sub("", m.group(2))
            if _BADBLOCK.search(tail) or _OPENBLOCK.search(tail):
                bad.append(c)
    return bad

def control_wiring(c=None):
    """For each form, every opt/chk/cmd/fra/lbl control referenced in its code
    must appear in that form's ControlsForForm Array(...).  Returns
    {form: [missing,...]} for forms with gaps."""
    c = content if c is None else c
    PRE = ("opt", "chk", "cmd", "fra", "lbl")
    SUF = ("_Click", "_Change", "_Enter", "_Exit", "_BeforeUpdate", "_AfterUpdate")
    gaps = {}
    for b in builder_names(c):
        if not b.startswith("frm"):
            continue
        form = b[:-5] if b.endswith("_Code") else b
        mm = re.search(r'Case "' + re.escape(form) +
                       r'"\r\n\s+ControlsForForm = Array\((.*?)\)\r\n', c, re.DOTALL)
        decl = set(re.findall(r'"(\w+)"', mm.group(1))) if mm else set()
        code = extract_code(b, c) or ""
        used = set()
        for x in re.finditer(r"\b((?:" + "|".join(PRE) + r")[A-Z]\w*)", code):
            nm = x.group(1)
            for sf in SUF:
                if nm.endswith(sf):
                    nm = nm[:-len(sf)]; break
            used.add(nm)
        missing = sorted(used - decl)
        if missing:
            gaps[form] = missing
    return gaps

def unterminated_strings(c=None):
    """Line numbers (1-based) with an odd number of quotes, excluding comments."""
    c = content if c is None else c
    return [n for n, l in enumerate(c.split("\r\n"), 1)
            if l.count('"') % 2 == 1 and not l.strip().startswith("'")]

_KNOWN_TOKENS = ("vbCrLf", "vbNewLine", "vbCr", "vbLf", "vbTab")

def unhandled_concat_tokens(c=None):
    """Builders must be pure literal/vb-const concatenation for extraction to be
    lossless.  Return {builder: [token,...]} for any accumulator-assignment token
    that is neither a "string literal" nor a known vb constant (e.g. Chr(34), a
    variable).  Uses the same accumulator-detection logic as extract_code so that
    non-standard accumulators (e.g. checklistText) are covered.
    Empty result == extraction is faithful for every builder."""
    c = content if c is None else c
    bad = {}
    for nm in builder_names(c):
        start = c.find(f"Private Function {nm}() As String")
        end = c.find("\r\nEnd Function", start)
        body = c[start:end]
        # detect accumulator the same way extract_code does
        acc = "s"
        for ln in body.split("\r\n"):
            st = ln.strip()
            m = re.match(r"(\w+)\s*=\s*(\w+)\s*$", st)
            if m and m.group(1) == nm:
                acc = m.group(2)
                break
        prefix = acc + " = " + acc + " & "
        for ln in body.split("\r\n"):
            st = ln.strip()
            if not st.startswith(prefix):
                continue
            expr = st[len(prefix):]
            n = len(expr); buf = ""; inq = False; i = 0; toks = []
            while i < n:
                ch = expr[i]
                if ch == '"': inq = not inq; buf += ch; i += 1
                elif ch == "&" and not inq: toks.append(buf); buf = ""; i += 1
                else: buf += ch; i += 1
            toks.append(buf)
            for t in toks:
                t = t.strip()
                if not t: continue
                if t.startswith('"') and t.endswith('"'): continue
                if t in _KNOWN_TOKENS: continue
                bad.setdefault(nm, []).append(t)
    return bad

def round_trip_ok(c=None, name=None):
    """Build fidelity: re-extracting the EMITTED code from a rebuilt builder must
    equal the original emitted code.  (Builder *source* may differ cosmetically --
    e.g. `& vbCrLf & vbCrLf` vs two lines -- and that is fine; what must be stable
    is the VBA the builder produces.)  Returns (ok, [failed_names])."""
    c = content if c is None else c
    names = [name] if name else builder_names(c)
    failed = []
    for nm in names:
        code = extract_code(nm, c)
        rebuilt = vba_to_builder(code, nm)
        if extract_code(nm, rebuilt) != code:
            failed.append(nm)
    return (not failed, failed)

def outer_module(c=None):
    """The host/installer code = everything minus the builder string lines.

    Uses accumulator-aware filtering: for each builder, detects its accumulator
    variable (same logic as extract_code) and filters out lines of the form
    `<acc> = <acc> & ...`.  This prevents non-standard accumulators like
    `checklistText = checklistText & ...` from leaking into the OUTER balance
    and if-trap checks.
    """
    c = content if c is None else c
    # Build the set of stripped prefixes to exclude, one per builder.
    exclude_prefixes = set()
    for name in builder_names(c):
        start = c.find(f"Private Function {name}() As String")
        if start < 0:
            continue
        end = c.find("\r\nEnd Function", start)
        body = c[start:end]
        acc = "s"
        for ln in body.split("\r\n"):
            st = ln.strip()
            m = re.match(r"(\w+)\s*=\s*(\w+)\s*$", st)
            if m and m.group(1) == name:
                acc = m.group(2)
                break
        exclude_prefixes.add(acc + " = " + acc + " & ")
    def _keep(line):
        stripped = line.strip()
        return not any(stripped.startswith(p) for p in exclude_prefixes)
    return "\n".join(l for l in c.split("\r\n") if _keep(l))

# ---------------------------------------------------------------- build gate
def run_all(c=None):
    """Run every gate. Returns (ok, report_lines)."""
    c = content if c is None else c
    rpt = []; ok = True

    # balance: outer module + each builder's injected code
    targets = [("OUTER", outer_module(c))] + [(b, extract_code(b, c)) for b in builder_names(c)]
    for label, code in targets:
        d, mn = balance(code)
        bad = [k for k in d if d[k] != 0 or mn[k] < 0]
        if bad:
            ok = False
            rpt.append(f"BALANCE  {label}: " + ", ".join(f"{k}={d[k]}(min {mn[k]})" for k in bad))

    # single-line-If traps across all builders
    traps = 0
    for b in builder_names(c):
        for line in single_line_if_traps(extract_code(b, c) or ""):
            traps += 1; ok = False
            rpt.append(f"IF-TRAP  {b}: {line[:70]}")
    # ...and the outer module
    for line in single_line_if_traps(outer_module(c)):
        traps += 1; ok = False
        rpt.append(f"IF-TRAP  OUTER: {line[:70]}")

    # control wiring
    for form, missing in control_wiring(c).items():
        ok = False
        rpt.append(f"WIRING   {form}: undeclared controls {missing}")

    # unterminated strings
    bad_lines = unterminated_strings(c)
    if bad_lines:
        ok = False
        rpt.append(f"STRINGS  odd-quote lines: {bad_lines[:8]}{' ...' if len(bad_lines) > 8 else ''}")

    # concat-token coverage (extraction faithfulness)
    bad_tok = unhandled_concat_tokens(c)
    if bad_tok:
        ok = False
        for nm, toks in bad_tok.items():
            rpt.append(f"TOKENS   {nm}: non-literal concat tokens {sorted(set(toks))}")

    # round-trip (emitted-code fidelity)
    rt_ok, failed = round_trip_ok(c)
    if not rt_ok:
        ok = False
        rpt.append(f"ROUNDTRIP failed for: {failed}")

    if ok:
        rpt.append(f"OK  {len(builder_names(c))} builders | balance, if-traps, "
                   f"wiring, strings, round-trip all clean")
    return ok, rpt

# ---------------------------------------------------------------- CLI
def _main(argv):
    global content
    args = [a for a in argv if not a.startswith("--")]
    fileopt = [a for a in argv if a.startswith("--file=")]
    if fileopt:
        content = load(fileopt[0].split("=", 1)[1])
    cmd = args[0] if args else "validate"

    if cmd == "validate":
        ok, rpt = run_all()
        print("\n".join(rpt))
        return 0 if ok else 1
    if cmd == "list":
        print("\n".join(builder_names()))
        return 0
    if cmd == "show":
        if len(args) < 2:
            print("usage: show <builderName>"); return 2
        print(extract_code(args[1]))
        return 0
    print(__doc__)
    return 0

if __name__ == "__main__":
    try:
        sys.exit(_main(sys.argv[1:]))
    except BrokenPipeError:
        # downstream closed the pipe (e.g. `| head`); not an error
        pass
