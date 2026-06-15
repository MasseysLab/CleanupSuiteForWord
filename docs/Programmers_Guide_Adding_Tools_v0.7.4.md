# CleanupSuite Programmer's Guide - Adding Tools

**Target release:** 0.7.4 documentation add-on
**Prepared:** 2026-06-14
**Audience:** Codex, Chris, and future maintainers

> Documentation basis: prepared for CleanupSuite 0.7.4 and reviewed against the local locked 0.7.4 project before commit. Screenshots are intentionally deferred.


## Plain-English overview

CleanupSuite is source-first. You do not edit the generated `VBA_Cleanup_tool.txt` by hand. You edit readable source files under `src/`, then run the assembler.

The normal loop is:

```text
edit source -> run python assemble.py -> test -> sync Word document/template -> commit
```

## Project pieces

Common files and folders:

```text
src/manifest.txt                       assembly order
src/installer/installer.bas             installer, controls, wiring, layout
src/installer/checklist.bas             manual install checklist text
src/modules/modCleanupHelpers.bas       shared safety, preview, settings helpers
src/modules/modCleanupLauncher.bas      version and launcher entry points
src/forms/frmCleanupSuiteLauncher.bas   main menu code
src/forms/frmPreviewActions.bas         preview action bar code
src/forms/frm*.bas                      individual tool forms
src/ribbon/ribbon.xml                   ribbon definition
src/ribbon/inject.bas                   ribbon injection helper
VBA_Cleanup_tool.txt                    generated distributable
assemble.py                             source-to-distributable builder
vbaeval.py                              validation/build-gate checks
tests/                                  Python regression tests
```

## Before adding a tool

Answer these questions first:

1. What exact document problem does the tool fix?
2. Is it safe enough for whole-document use?
3. Does it need selection scope?
4. Can it support preview before apply?
5. What should the user see after preview?
6. What should happen if the document is blank or protected?
7. What should Undo do?
8. What test proves the tool is wired correctly?

If the tool can change or delete large amounts of content, design it as preview-first.

## Simple checklist for adding a new tool

### 1. Create the form source file

Add a new file under `src/forms/`, for example:

```text
src/forms/frmMyNewCleanup.bas
```

Give it a builder function name that matches the manifest pattern, for example:

```text
frmMyNewCleanup_Code
```

### 2. Add the tool to the manifest

Edit `src/manifest.txt` and add a builder entry near similar tools:

```text
BUILDER  forms/frmMyNewCleanup.bas  frmMyNewCleanup_Code
```

The manifest controls the order that `assemble.py` uses when generating `VBA_Cleanup_tool.txt`.

### 3. Add launcher wiring

In `src/forms/frmCleanupSuiteLauncher.bas`, add a button click handler that opens the form through the shared launcher lifecycle helper:

```vba
Private Sub cmdMyNewCleanup_Click(): OpenCleanupTool "frmMyNewCleanup": End Sub
```

Do not use `Me.Hide: frmMyNewCleanup.Show` for new tools. In 0.7.4 the launcher creates a fresh form instance with `VBA.UserForms.Add(...)` so a tool can be opened, closed, and reopened without the MSForms â€œform already displayedâ€ runtime error.

Also add Help text if the launcher includes a Help button for the tool.

### 4. Add installer/control wiring

In `src/installer/installer.bas`, add the form to the install list, control list, layout routine, and any friendly captions/tooltips required by the current installer pattern.

A form is not fully installed just because the source file exists. The installer must know how to create the form and its controls.

### 5. Use shared safety helpers

Most tools should use existing helpers instead of inventing their own safety flow.

Common helpers include:

```vba
GuardBeforeCleanup "Tool Name"
MarkCleanupStart "Tool Name"
MarkCleanupEnd
RemoveAllHighlighting ActiveDocument.Content
DocumentHasVisibleContent()
ShowPreviewActions Me, "Tool Name", summaryText
ReturnToMainAfterToolClose
```

### 6. Support preview and apply

For preview-style tools, use a clear split:

```text
Preview path = highlight/report only
Apply path = make actual changes
```

The Preview Actions bar expects source forms to provide methods such as:

```vba
PreviewFromPanel
RunAfterPreview
```

Use the existing tools as patterns.

### 7. Add tests

Add or update tests under `tests/`.

Good tests check that:

- the manifest includes the new builder
- the launcher opens the new form
- the installer creates the expected controls
- preview wording is present
- safety guards are present
- generated bundle includes the new code

### 8. Build and test

Run:

```text
python assemble.py
python -m unittest discover -s tests
```

Then sync and smoke-test the `.docm`/template in Word.

## Simple coding rules

- Keep generated text out of hand edits.
- Do not edit `VBA_Cleanup_tool.txt` manually.
- Use the shared helpers for preview, safety, settings, and return-to-main behavior.
- Keep user-facing text plain and specific.
- Prefer preview before destructive edits.
- Add tests before calling a new tool complete.

## Minimum user-facing help section

Every tool should be explainable in this structure:

```text
What it fixes
Use it when
Preview guidance
Best workflow
Caution
```

That structure also makes the User Manual easier to maintain.

## Advanced section for experienced developers

### Builder functions

The generated distributable is assembled from builder functions. A `BUILDER` manifest entry converts a readable source file into a VBA function that returns text. This allows form code to live as source while still shipping inside one pasteable installer file.

### Installer responsibility

The installer is responsible for:

- creating/removing VBA components
- naming forms correctly
- creating controls
- applying captions, fonts, sizes, positions, and tooltips
- wiring modules and forms into the active Word project
- reporting install problems

If a project only shows `UserForm1` after install, the install probably failed during form creation or naming. Do not fix that by changing launcher code. Fix the form installation path.

### Preview contract

The preview system should keep these responsibilities separate:

- tool form chooses settings
- preview run highlights/reports candidates
- Preview Actions bar lets the user apply, reconfigure, or toggle preview state
- apply run performs changes using the same settings
- shared helpers manage preview panel position and cleanup state

### Safety state

Cleanup operations should mark start/end state where appropriate so interrupted operations can be detected later. Tools should also respect protected documents, blank documents, and the macro-host document.

### Tests as documentation

Tests are not just error checks. They describe the expected architecture. When a new tool changes installer wiring, preview behavior, form controls, or the launcher, update tests to document that new contract.

### Release discipline

For documentation-only releases, do not change tool behavior. For code releases, update `SUITE_VERSION`, rebuild generated bundles, sync the practice bundle, sync embedded `.docm` files, and record the result in `AI_HANDOFF.md`.

## Advanced TODO for future 1.0.0

The long-term distribution model should avoid modifying `Normal.dotm`. The likely 1.0.0 target is a separate global template/add-in:

```text
CleanupSuite.dotm
%APPDATA%\Microsoft\Word\Startup
```

Do not convert to that model until the project is ready for the global add-in/template milestone.
