# CleanupSuite Programmer's Guide For Adding Tools

**Target release:** 0.9.0 Alpha (`v0.9.0-alpha`)
**Prepared:** 2026-07-14
**Audience:** contributors and future maintainers

Copyright (c) 2026 MasseysLab. Software and documentation are licensed under the MIT License.


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
src/installer/installer.bas            installer, controls, wiring, layout
src/installer/checklist.bas            manual install checklist text
src/modules/modCleanupHelpers.bas      shared safety, preview, settings, layout helpers
src/modules/modCleanupLauncher.bas     version and launcher entry points
src/forms/frmCleanupSuiteLauncher.bas  main menu code
src/forms/frmPreviewActions.bas        preview action window code
src/forms/frm*.bas                     individual tool forms
src/ribbon/ribbon.xml                  ribbon definition
src/ribbon/inject.bas                  ribbon injection helper
VBA_Cleanup_tool.txt                   generated distributable
assemble.py                            source-to-distributable builder
vbaeval.py                             validation/build-gate checks
tests/                                 Python regression tests
docs/Tool-Form Rules.md                current shared tool-form layout rules
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

In `src/forms/frmCleanupSuiteLauncher.bas`, add a button click handler that opens the form:

```vba
Private Sub cmdMyNewCleanup_Click(): OpenCleanupTool "frmMyNewCleanup": End Sub
```

Do not use `Me.Hide: frmMyNewCleanup.Show` for new tools. The launcher opens a fresh form instance through `OpenCleanupTool`, which avoids the old "form already displayed" problem and keeps return-to-main behavior consistent. Most forms use `VBA.UserForms.Add(...)`; add-in-sensitive advanced forms such as `frmMetaDataSuite` use a compile-time `New frmMetaDataSuite` branch in `CreateCleanupToolForm` because string-created forms can fail when the same code is hosted in a Startup DOTM.

If the launcher includes a Help button for the tool, wire that help from the main menu, not from the tool form itself.

If you add a new launcher row after older `.docm` files already exist, do not assume the embedded launcher designer will automatically match the source tree. New launcher controls may be added later by the sync path and can show default MSForms appearance until the launcher normalizes them at runtime.

Practical rule:

- if a new launcher row must visually match an existing row, copy only safe runtime members such as `Caption`, `Left`, `Top`, `Width`, `BackColor`, and `ForeColor`
- if the row depends on visual emphasis, normalize text roles explicitly at runtime too: for example, help/tool buttons bold, explanation label normal
- do not casually copy every visible property from one launcher button to another in form code-behind
- in particular, avoid assuming a member such as `SpecialEffect` is safe to copy in early-bound UserForm code just because the installer can set it through generic late-bound control objects

### 4. Add installer/control wiring

In `src/installer/installer.bas`, add the form to the install list, control list, layout routine, and any friendly captions/tooltips required by the current installer pattern.

A form is not fully installed just because the source file exists. The installer must know how to create the form and its controls.

If the new tool adds launcher controls, also check the `.docm` sync path in `scripts/sync_docm_code_only.ps1`. Older workspace documents may need explicit creation of the new controls before the refreshed code can position or caption them.

### 5. Use shared safety helpers

Most tools should use existing helpers instead of inventing their own safety flow.

Common helpers include:

```vba
GuardBeforeCleanup "Tool Name"
MarkCleanupStart "Tool Name"
MarkCleanupEnd
BeginCleanupProgress "Tool Name", "Starting"
UpdateCleanupProgress "Readable phase", completedItems, totalItems
EndCleanupProgress
RemoveAllHighlighting ActiveDocument.Content
DocumentHasVisibleContent()
ParagraphContainedInRange paragraphItem, scopeRange
ParagraphOverlapsRange paragraphItem, scopeRange
ParagraphOverlapsRangeOutsideTable paragraphItem, scopeRange
ParagraphBodyText(paragraphItem)
ReplaceParagraphBodyText paragraphItem, replacementText
Dim previewRows As Collection
Set previewRows = NewPreviewSummaryRows()
AddPreviewSummaryRow previewRows, "Human item label", itemCount
ShowPreviewActionsSummary Me, "Tool Name", previewRows
ReturnToMainAfterToolClose Me, Cancel, CloseMode
LayoutCleanupToolForm Me
```

Use the existing tools as patterns before inventing a new local helper.

The paragraph helpers deliberately preserve different scope meanings. Use `ParagraphContainedInRange` when a partially selected paragraph must not be changed. Use `ParagraphOverlapsRange` when paragraph-level formatting or cleanup should include a paragraph touched by the scope. Use `ParagraphOverlapsRangeOutsideTable` only for tools whose contract excludes table paragraphs. `ParagraphBodyText` and `ReplaceParagraphBodyText` preserve both ordinary paragraph marks and table-cell markers; do not recreate local versions that trim `vbCr` or `Chr$(7)` independently.

All user-visible progress belongs to the shared progress service. Start it once, update it at meaningful intervals rather than for every character, and guarantee `EndCleanupProgress`, `MarkCleanupEnd`, or `RestoreCleanupSuiteTransientState` on every normal, cancel, and error exit. Tool code must not write directly to `Application.StatusBar`.

### 6. Follow the shared tool-form rules

Most tool forms now follow the shared guided layout. Read `docs/Tool-Form Rules.md` before changing tool-form layout behavior.

Important current rules:

- the visible title belongs in the body, not the native form title bar
- the centered bold line is the instruction, not a repeated title
- Help stays on the main menu, not inside tool forms
- Preview spans the full row
- Reset is a small utility in the upper-right title row
- Preview is the visible bottom action; Apply appears only in Preview Actions
- Custom sections, divider lines, Select All, and Deselect All appear only when the current tool state calls for them
- Header / Footer Standardizer has special layout handling
- every form must fit a 1024 x 768 **pixel** display; do not use those pixel numbers as VBA form points
- the approved maximum outer UserForm envelope is `768 x 510` VBA form points, matching the main menu; individual forms should be smaller when their content permits it

If a form is meant to behave like the other guided tools, it should call the shared layout path instead of hand-tuning its own spacing.

MSForms positions and sizes UserForms in points, while the target monitor requirement is expressed in pixels. Keep those units separate. Leave room inside the `768 x 510` point envelope for the native title bar and borders, and verify the bottommost control against `InsideHeight`. The Safe Metadata Editor is a current dense-layout reference: it uses `760 x 510` form points and keeps all controls visible by removing unused vertical gaps. Always perform a live Word check at the target 1024 x 768 pixel display size or an equivalent scaled capture; source constants alone do not prove that text and bottom actions are visible.

### 7. Support preview and apply

For preview-style tools, use a clear split:

```text
Preview path = highlight/report only
Apply path = make actual changes
```

The Preview Actions window expects source forms to provide methods such as:

```vba
PreviewFromPanel
RunAfterPreview
```

Use the existing tools as patterns.

### 8. Add tests

Add or update tests under `tests/`.

Good tests check that:

- the manifest includes the new builder
- the launcher opens the new form
- the installer creates the expected controls
- the `.docm` sync script knows about any new launcher or form controls that must exist in older embedded projects
- any launcher-row runtime refresh code sets required captions and only uses safe visual-property normalization
- any runtime launcher refresh also restores the intended bold/normal text hierarchy and any non-default button presentation that the row depends on
- preview wording is present
- safety guards are present
- generated bundle includes the new code
- shared guided-layout expectations still hold if the tool uses the common layout path

### 9. Build and test

Run:

```text
python assemble.py
python -m unittest discover -s tests
```

Then sync and smoke-test the `.docm`/template in Word.

If generated root artifacts change, update the matching practice copy in `Practice - Try CleanupSuite Here` in the same pass.

For launcher changes, do one extra real Word check on an already-existing `.docm`, not only on rebuilt source files. That is the fastest way to catch:

- missing newly added controls
- default-looking buttons caused by stale embedded designers
- help buttons missing their `?` caption
- stale font-weight or caption presentation, such as a tool name losing bold emphasis or a footer button falling back to a plain default look
- VBA compile errors caused by unsafe copied control members

## Simple coding rules

- Keep generated text out of hand edits.
- Do not edit `VBA_Cleanup_tool.txt` manually.
- Use the shared helpers for preview, safety, settings, layout, and return-to-main behavior.
- Prefer descriptive procedure and variable names, one action per line, and short comments that explain the reason for a non-obvious rule.
- Add a shared helper only when current code uses it immediately; replace the duplicated implementations in the same change.
- Do not retain compatibility wrappers, unused generalized frameworks, or dead helpers. A small tool-local helper is better when only one tool needs the behavior.
- Keep Preview analysis inside the Preview path. Apply should not first count or highlight the same candidates unless that work is required to make the change safely.
- Keep user-facing text plain and specific.
- Prefer preview before destructive edits.
- Add tests before calling a new tool complete.
- When touching guided forms, stay consistent with `docs/Tool-Form Rules.md`.

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

### Guided layout contract

The shared guided layout should keep these responsibilities separate:

- the tool form owns its settings and behavior
- `LayoutCleanupToolForm` owns the common arrangement and compact spacing
- `ApplyMSFormTitleStrategy` controls whether the body owns the visible title
- the main menu owns per-tool Help buttons
- special-case layout should stay narrow and justified, not become the default pattern

The global size contract also applies to special-case forms: no UserForm may exceed the main menu's `768 x 510` VBA point envelope for the supported 1024 x 768 pixel screen. If content does not fit, reduce nonfunctional whitespace, pair related fields, wrap explanatory text deliberately, or use a compact scrollable content region. Do not shrink button text until it clips or treat 1024 x 768 pixels as 1024 x 768 VBA points.

The guided info box (`LayoutGuidedInfoBox`) uses a two-column design: a left column for a concise display name and a right column for plain-language advisory text. Display names come from `GuidedInfoDisplayName()` and advisory text from `GuidedInfoAdvisory()`, both keyed by the control name. Row heights are dynamic. When adding a new tool, add entries for any new controls to both lookup functions in `modCleanupHelpers.bas`. Controls without a specific entry fall back to `GuidedGenericAdvisory()`.

### Preview contract

The preview system should keep these responsibilities separate:

- tool form chooses settings
- preview run highlights/reports candidates
- Preview Actions window lets the user apply, reconfigure, toggle preview state, or turn preview pages
- apply run performs changes using the same settings
- shared helpers manage preview panel position and cleanup state

### Safety state

Cleanup operations should mark start/end state where appropriate so interrupted operations can be detected later. Tools should also respect protected documents, blank documents, and the macro-host document.

### Tests as documentation

Tests are not just error checks. They describe the expected architecture. When a new tool changes installer wiring, preview behavior, form controls, launcher behavior, or guided layout assumptions, update tests to document that contract.

When adding or renaming a control on an existing UserForm, update every creation path that may touch that form:

- the form source in `src/forms/`
- `ControlsForForm()` and any form-specific layout routine in `src/installer/installer.bas`
- `ControlCaptionText()` in `src/installer/installer.bas`
- guided display/advisory text in `src/modules/modCleanupHelpers.bas` when the control appears in the guided info area
- the relevant sync/installer regression tests

This matters because older embedded `.docm` files may not already contain newly added controls. Runtime code should only reference a control after the installer/sync path knows how to create it.

### Release discipline

For documentation-only releases, do not change tool behavior. For code releases, update `SUITE_VERSION`, rebuild generated bundles, sync the practice bundle, sync embedded `.docm` files, and record the verification result in the release notes or pull request.

## Advanced TODO for future 1.0.0

The long-term distribution model should avoid modifying `Normal.dotm`. The likely `1.0.0` target is a separate global template/add-in:

```text
CleanupSuite.dotm
%APPDATA%\Microsoft\Word\Startup
```

Do not convert to that model until the project is ready for the global add-in/template milestone.
