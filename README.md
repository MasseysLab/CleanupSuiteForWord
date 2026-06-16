<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 Christopher Travis Massey — see LICENSE for full terms. -->

# VBA Cleanup Suite — Build Pipeline

## Overview

The distributable (`VBA_Cleanup_tool.txt`) is a single CRLF file the user pastes
into a VBA module and runs.  This pipeline gives you a proper authoring workflow:
readable per-component source files, a build step that assembles them, and a
validation gate that fails loudly on errors.

```
src/                   ← authorable source tree
  manifest.txt         ← assembly order (generated once by explode, then maintained)
  installer/
    installer.bas      ← host module: install/uninstall/wiring subs (one VBComponent)
    checklist.bas      ← ManualChecklistText content (plain text, no builder strings)
  modules/
    modCleanupHelpers.bas
    modCleanupLauncher.bas
  forms/
    frmPunctuationCleanup.bas
    frmUnicodeCleanup.bas
    ... (one .bas per UserForm)
    frmCleanupSuiteLauncher.bas
  ribbon/
    ribbon.xml         ← plain XML injected at ribbon install time
    inject.bas         ← InjectRibbonXml sub + Wait helper

VBA_Cleanup_tool.txt   ← assembled distributable (do not edit by hand)
explode.py             ← splits an existing .txt into src/
assemble.py            ← rebuilds the .txt from src/ + runs build-gate
vbaeval.py             ← validation library + CLI (used by assemble.py)
```

---

## Normal workflow

**Edit a form or module:**

1. Open its `.bas` file in `src/forms/` or `src/modules/` and edit.
2. Run `python assemble.py`.  If validation passes, `VBA_Cleanup_tool.txt` is updated and
   `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt` is refreshed to match.
3. Paste the updated `.txt` into Word and test.

**Edit the installer (wiring, control layout, etc.):**

1. Edit `src/installer/installer.bas`.
2. Run `python assemble.py`.

**Add a new tool** (see `CONTRIBUTING.md` for the full wiring checklist):

1. Create `src/forms/frmMyTool.bas` with the form's VBA.
2. Wire the remaining touch-points in `installer.bas` and `frmCleanupSuiteLauncher.bas`.
3. Add the new builder entry to `manifest.txt` (copy the pattern of an adjacent form).
4. Follow the shared two-column guided layout rule for generic tool choices and Custom choices, keep the forms intro-first, and make sure single leftover choices are centered.
5. Run `python assemble.py`.
6. If a repo artifact has a matching practice copy, sync that practice file in the same pass.

---

## Versioning

The current release is `0.8.0`, which carries forward the locked `0.7.5` baseline and the final pre-`0.8` polish pass into the next programming milestone.

See [`VERSIONING.md`](VERSIONING.md) for the `0.5.0` through `1.0.0` roadmap and the rule for when `SUITE_VERSION` should change.

---

## Scripts

### `explode.py` — one-time migration from an existing `.txt`

Splits a distributable `.txt` into the `src/` tree and generates `manifest.txt`.
Run this once when bootstrapping; after that, maintain the source files directly.

```
python explode.py [--file=VBA_Cleanup_tool.txt] [--src=src]
```

What it produces:

- `installer/installer.bas` — the outer host module verbatim (lines 1 through
  end of `ControlTypeByName`, the last helper function before the builders).
- `installer/checklist.bas` — emitted text of `ManualChecklistText` (plain, readable).
- `modules/*.bas` and `forms/*.bas` — emitted VBA for each injected component.
- `ribbon/ribbon.xml` — plain XML from the `RIBBON_XML` builder.
- `ribbon/inject.bas` — verbatim ribbon module sections.
- `manifest.txt` — the assembly-order file consumed by `assemble.py`.

All `.bas` and `.xml` source files use **LF line endings**.  `assemble.py` converts
to CRLF when building the distributable.

### `assemble.py` — rebuild the distributable

Reads `manifest.txt`, assembles chunks in order, validates, and writes the output.

```
python assemble.py [--src=src] [--out=VBA_Cleanup_tool.txt] [--no-validate]
```

Exits 0 on success, exits 1 if any validator fails.  The output file is **not
written** when validation fails.

### `vbaeval.py` — validation library + CLI

Run the build-gate validators directly against any `.txt`:

```
python vbaeval.py validate [--file=VBA_Cleanup_tool.txt]
python vbaeval.py list
python vbaeval.py show frmSpacingCleanup_Code
```

`validate` exits 0 if clean, 1 on any failure.  Wire it into CI or a pre-commit
hook to catch problems before pushing:

```yaml
# .github/workflows/build.yml (example)
- run: python assemble.py --src=src --out=VBA_Cleanup_tool.txt
```

---

## `manifest.txt` format

The manifest is the assembly order for `assemble.py`.  Each non-blank,
non-comment line has the form:

```
TYPE  path  [func_name]
```

| Type | Meaning |
|------|---------|
| `VERBATIM` | Read the file, convert LF→CRLF, paste into distributable as-is. |
| `BUILDER` | Read the file (plain VBA), run `vba_to_builder` to produce the `Private Function func_name() As String` block. |
| `XML_BUILDER` | Same as `BUILDER` but source is plain XML (used for `ribbon.xml`). |
| `SEP` | Literal CRLF separator block (base64-encoded), pasted verbatim. Carries the comment lines between builders. |

`SEP` entries are generated by `explode.py` and should not be edited by hand.
To add a new tool, copy the `SEP` + `BUILDER` pair from an adjacent entry and
update the path and function name.

---

## Build-gate validators (run automatically by `assemble.py`)

All five run on every `assemble.py` invocation:

| Check | What it catches |
|-------|----------------|
| **Balance** | `Sub`/`Function`/`With`/`If`/`For`/`Do`/`Select` nesting doesn't return to 0, or goes negative. Uses a string-aware tokeniser that ignores `:` and `"` inside literals. |
| **If-traps** | Single-line `If x Then <block>` where the block contains `With`, `For`, `Do`, `End If`, etc. These compile-fail in VBA. |
| **Control wiring** | Every `opt`/`chk`/`cmd`/`fra`/`lbl` control referenced in a form's code is declared in that form's `ControlsForForm` case. |
| **Unterminated strings** | Lines with an odd quote count (excluding comment lines). |
| **Round-trip** | Re-extracting the emitted code from a rebuilt builder is byte-identical to the original extract. Catches non-literal concatenation in builder source. |

Run validators without assembling:

```
python vbaeval.py validate
```

---

## Line-ending discipline

| File | Endings | Why |
|------|---------|-----|
| Source `.bas` / `.xml` in `src/` | **LF** | Readable in any editor, normal for version control. |
| `VBA_Cleanup_tool.txt` (distributable) | **CRLF** | Required: VBA editor on Windows expects CRLF; the self-injector reads the file with `newline=''`. |
| `vbaeval.py`, `explode.py`, `assemble.py` | LF | Normal Python scripts. |

`assemble.py` handles the LF→CRLF conversion automatically.  Never edit the
distributable `.txt` directly — always edit source and reassemble.

---

## What the diff means after the first explode/assemble

The assembled distributable is **not byte-identical** to the original `.txt` in
two intentional ways:

1. **Bare CRs in `IsSuiteComponent`** — the original had `\r\r\n` line-continuation
   artifacts (8 occurrences) from an earlier editor session.  The assembled file
   normalises these to `\r\n`.  VBA ignores the difference.

2. **`ManualChecklistText` builder** — the original used a `checklistText`-named
   accumulator; the assembled version uses the canonical `s` accumulator.
   The emitted string at runtime is byte-identical.

All 23 builders emit identical VBA.  The distributable compiles and runs correctly.

---

## License

MIT License — see [`LICENSE`](LICENSE) for full terms.  
Copyright (c) 2026 Christopher Travis Massey
