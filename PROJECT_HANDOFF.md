# VBA Cleanup Suite — Project Handoff

## What this project is

A self-installing collection of 19 document-cleanup tools for Microsoft Word,
shipped as a single CRLF `.txt` file (`VBA_Cleanup_tool.txt`) the user pastes
into a VBA module and runs.  At runtime it self-injects 20 UserForms + 2 modules
into the VBA project via VBIDE.

A full authoring/build pipeline has been built and is working cleanly.

---

## Current state (as of this handoff)

- Pipeline is complete and verified end-to-end on the real file.
- `explode.py` and `assemble.py` both ran clean on the developer's machine.
- All 5 build-gate validators pass: balance, if-traps, control wiring,
  unterminated strings, round-trip identity.
- All 23 builders emit identical VBA between original and assembled output.
- All files carry MIT license headers.
- `CONTRIBUTING.md` documents the 8-point wiring checklist for adding new tools.

---

## Complete working directory

```
CleanupSuite/
│
├── LICENSE
├── README.md
├── CONTRIBUTING.md
│
├── VBA_Cleanup_tool.txt        ← the distributable (do not edit by hand)
│
├── vbaeval.py                  ← validation library + CLI
├── assemble.py                 ← rebuilds .txt from src/ + runs build gate
├── explode.py                  ← one-time bootstrap (keep, don't delete)
│
└── src/
    ├── manifest.txt
    ├── installer/
    │   ├── installer.bas       ← host module: install/uninstall/wiring
    │   └── checklist.bas       ← ManualChecklistText content
    ├── modules/
    │   ├── modCleanupHelpers.bas
    │   └── modCleanupLauncher.bas
    ├── forms/
    │   ├── frmPunctuationCleanup.bas
    │   ├── frmUnicodeCleanup.bas
    │   ├── frmSpacingCleanup.bas
    │   ├── frmCapitalizationCleanup.bas
    │   ├── frmListCleanup.bas
    │   ├── frmParagraphCleanup.bas
    │   ├── frmDuplicateDetector.bas
    │   ├── frmFontNormalizer.bas
    │   ├── frmTableCleaner.bas
    │   ├── frmBreakNormalizer.bas
    │   ├── frmDocumentTrim.bas
    │   ├── frmFormattingStripper.bas
    │   ├── frmHyperlinkRemover.bas
    │   ├── frmSoftReturnConverter.bas
    │   ├── frmMetadataScrubber.bas
    │   ├── frmStyleCleanup.bas
    │   ├── frmFootnoteRemover.bas
    │   ├── frmHeaderFooterStandardizer.bas
    │   ├── frmObjectRemover.bas
    │   └── frmCleanupSuiteLauncher.bas
    └── ribbon/
        ├── ribbon.xml
        └── inject.bas
```

11 files at the root, 28 files inside `src/`.  Total: 39 files.

---

## Daily workflow

**Edit a form or module:**
1. Open its `.bas` file in `src/forms/` or `src/modules/` and edit.
2. Run `python assemble.py` — validates and rebuilds `VBA_Cleanup_tool.txt`.
3. Paste the updated `.txt` into Word and test.

**Edit the installer (wiring, control layout, etc.):**
1. Edit `src/installer/installer.bas`.
2. Run `python assemble.py`.

**Add a new tool:**
See `CONTRIBUTING.md` for the full 8-point wiring checklist.
In brief:
1. Create `src/forms/frmMyTool.bas` with the form's VBA.
2. Wire the 8 integration points in `installer.bas` and `frmCleanupSuiteLauncher.bas`.
3. Add a `SEP` + `BUILDER` entry to `src/manifest.txt` (copy an adjacent form's pattern).
4. Run `python assemble.py` — fix any validator failures before proceeding.

**Validate without assembling:**
```
python vbaeval.py validate
python vbaeval.py list
python vbaeval.py show frmSpacingCleanup_Code
```

---

## Key architecture facts

- Builders are `Private Function frmXxx_Code() As String`, placed before
  `frmCleanupSuiteLauncher_Code` in the distributable.
- `vbaeval.py` extract_code() is accumulator-aware — detects the accumulator
  variable from the return line rather than hard-coding `s`.  This handles
  `ManualChecklistText` which uses `checklistText` as its accumulator.
- `outer_module()` in `vbaeval.py` uses the same accumulator-aware filtering
  so non-`s` lines never leak into the OUTER balance/if-trap checks.
- The distributable uses CRLF throughout.  Source files in `src/` use LF.
  `assemble.py` handles the conversion automatically.
- `assemble.py` finds `vbaeval.py` by looking in the same directory as itself,
  then in `cwd`.  Keep both files at the project root.

---

## Known intentional diffs from the very first original .txt

After the first explode/assemble, the rebuilt file differs from the original
in two harmless ways (already normalised in the current file):
1. 8 bare CRs (`\r\r\n`) in `IsSuiteComponent` normalised to `\r\n`.
2. `ManualChecklistText` builder normalised to standard `s` accumulator.
All 23 builders emit identical VBA at runtime.

---

## What comes next (in order of priority)

### 1. Version control (do this first — 10 minutes)
Put the project into Git.  Every future edit becomes reversible.

```
git init
git add .
git commit -m "Initial commit — pipeline working, all validators clean"
```

Then create a repository on GitHub and push.  The project is MIT licensed
under the name Christopher Travis Massey, so it can be public immediately
if desired.

### 2. Test the suite in Word
Run `InstallCleanupSuite` on a real document and exercise all 19 tools.
This is the one thing the sandbox cannot do.  Confirm everything installs
and operates correctly before building further.

### 3. Add a new tool
The real test of the pipeline and `CONTRIBUTING.md`.  Pick a simple tool,
follow the 8-point checklist, run `assemble.py`, and confirm the new form
installs and works.  This will surface any gaps in `CONTRIBUTING.md`.

### 4. The .dotm repackaging (major future milestone)
Repackage as a prebuilt `.dotm` template with a real ribbon instead of
the paste-and-run delivery.  This is a separate, larger piece of work that
the current pipeline feeds directly into.  Tackle after steps 1–3 are solid.

---

## How to start the next session

Upload all seven root files:
- `VBA_Cleanup_tool.txt`
- `vbaeval.py`
- `assemble.py`
- `explode.py`
- `README.md`
- `CONTRIBUTING.md`
- `LICENSE`

Then open with:
> "This is the VBA Cleanup Suite project. The pipeline is built and working —
> explode/assemble runs clean on my machine. Here are all the project files.
> I want to [next thing]."

---

## License

MIT License — Copyright (c) 2026 Christopher Travis Massey
See `LICENSE` for full terms.
