<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 Christopher Travis Massey — see LICENSE for full terms. -->

# Contributing — How to Add a Tool to the Cleanup Suite

This guide explains how to add a new cleanup tool to the Professional Cleanup
Suite for Microsoft Word. Read it **in full before writing any VBA** — several
of the cautions below are non-obvious compile traps that will silently break the
installer if you miss them.

---

## 1. How the suite is built (the 60-second model)

The entire suite ships as a single `.txt` file that you paste into one standard
module and run once (`InstallCleanupSuite`). At install time it programmatically
creates the UserForms and helper modules inside the host document/template using
the VBA Extensibility (VBIDE) object model.

Because a plain `.txt` cannot contain real UserForm binaries, **each tool's form
code is stored as a string-building function** named `frmXxx_Code`. That function
returns the form's source as text; the installer injects it into a freshly created
UserForm. So when you "write a form," you are really writing VBA *inside a string*.

Key pieces:

- **`InstallCleanupSuite`** — the installer. Creates every form + module.
- **`frmXxx_Code()` builder functions** — one per form; emit the form's code as text.
- **`modCleanupHelpers`** — shared utilities every tool calls (scope, guards,
  reporting, highlight clearing).
- **`modCleanupLauncher`** — ribbon hook + `ShowCleanupSuiteLauncher`.
- **`frmCleanupSuiteLauncher`** — the menu form with a button per tool.
- **`ControlsForForm`** — declares which controls each form has.
- **`AttemptCreateControls` + `LayoutFormControls`** — create and auto-position
  controls at install time.

The ribbon has a **single** button that opens the launcher, so adding a tool does
**not** require any ribbon edits — the tool appears once it is wired into the
launcher.

---

## 2. Recommended workflow

Editing VBA *inside builder strings* by hand is error-prone (every `"` must be
doubled). Do this instead:

1. **Develop the form live.** In a test document's VBA editor, insert a real
   UserForm, add the controls, write the code, and get it to **Debug → Compile**
   and run correctly.
2. **Convert to a builder function** only once it works (see §4).
3. **Wire it into the eight integration points** (see §3), and add the new
   `SEP` + `BUILDER` entry to `src/manifest.txt` (copy the pattern of an
   adjacent form entry).
4. **Run `python assemble.py`** — it rebuilds `VBA_Cleanup_tool.txt` and runs
   all five build-gate validators.  Fix any failures before proceeding.
5. **Run the full test checklist** (see §8) on a throwaway copy.

---

## 3. The eight places a new tool must be wired

Adding a tool named, say, `frmWordCounter` means touching **eight** locations.
Miss one and you get either a missing form, a dead button, or a control that is
never created.

1. **Builder function** — add `Private Function frmWordCounter_Code() As String`
   (the form's code as a string). Place it alongside the other `frmXxx_Code`
   functions, before `frmCleanupSuiteLauncher_Code`.

2. **Installer — create the form.** In `InstallCleanupSuite`, add:
   ```vba
   CreateOrReplaceForm vbProj, "frmWordCounter", frmWordCounter_Code, installLog
   ```
   alongside the other `CreateOrReplaceForm` calls (before the launcher line).

3. **Installer — create its controls.** In `InstallCleanupSuite`, add:
   ```vba
   AttemptCreateControls vbProj, "frmWordCounter", uiFailures
   ```
   alongside the other `AttemptCreateControls` calls.

4. **`ControlsForForm`** — add a `Case` returning the control-name array
   **in visual top-to-bottom order** (the auto-layout uses this order):
   ```vba
   Case "frmWordCounter"
       ControlsForForm = Array("chkOptionA", "chkOptionB", _
                               "chkPreviewOnly", "cmdRun", "cmdHelp", _
                               "fraScopeSelection", "optScopeDocument", "optScopeSelection")
   ```

5. **Launcher — open handler.** In `frmCleanupSuiteLauncher_Code`, add:
   ```vba
   Private Sub cmdWordCounter_Click(): frmWordCounter.Show: End Sub
   ```

6. **Launcher — help handler.** In `frmCleanupSuiteLauncher_Code`, add a
   `cmdHelpWordCounter_Click()` sub that shows the tool's help text.

7. **Launcher control list.** In `ControlsForForm`'s
   `Case "frmCleanupSuiteLauncher"` array, add the tool button (`cmdWordCounter`)
   to the tool-button group and the help button (`cmdHelpWordCounter`) to the
   help-button group.

8. **Uninstaller component list.** In the host module's `IsSuiteComponent`
   function, add `"frmWordCounter"` to the `names = Array(...)` list. The
   developer tool `UninstallCleanupSuite` uses that list to know what to strip
   before a clean reinstall — a tool missing from it will be left behind.

(Optional: `ManualChecklistText` lists control names for a few forms as a manual
fallback when programmatic control creation is blocked. Adding your tool there is
nice-to-have, not required — most forms are not listed.)

---

## 4. Builder-function format (exact rules)

When converting your working form code to a `frmXxx_Code` function:

```vba
Private Function frmWordCounter_Code() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private Sub cmdRun_Click()" & vbCrLf
    s = s & "    MsgBox ""Hello""" & vbCrLf
    s = s & "End Sub"
    frmWordCounter_Code = s
End Function
```

Rules — **all four matter**:

- **Every `"` inside the emitted code becomes `""`.** `MsgBox "Hi"` is written as
  `s = s & "MsgBox ""Hi"""`. Apostrophes (`'`, for comments) are *not* doubled.
- **The last content line omits `& vbCrLf`.** Every line ends in `" & vbCrLf`
  *except the final one*, which ends in just `"`. (This matches the existing
  builders; a stray trailing newline is harmless but breaks byte-for-byte tooling.)
- **One statement per line.** Do not build a single emitted line with VBA line
  continuations (`_`). The `s = s & "..."` pattern is deliberately
  continuation-free so it can never hit VBA's ~24-continuations-per-statement cap.
- **Keep `ControlsForForm` array continuations modest.** If a control list is
  long, split it across a few `_` lines (the existing ones use 3–4) — well under
  the cap.

---

## 5. Control naming rules (the prefix *is* the type)

`AttemptCreateControls` chooses each control's type from its **name prefix** via
`ControlTypeByName`. Name controls correctly or they will be created as the wrong
type:

| Prefix | Control created      |
|--------|----------------------|
| `opt`  | OptionButton         |
| `chk`  | CheckBox             |
| `cmd`  | CommandButton        |
| `fra`  | Frame                |
| `lbl`  | Label                |
| (other)| TextBox              |

Conventions used across the suite, which you should follow for consistency and so
the auto-layout reads well:

- `cmdRun` — the action button. `cmdHelp` — the in-form help button.
- `chkPreviewOnly` — the preview toggle (almost every tool has one).
- Scope controls: `fraScopeSelection`, `optScopeDocument`, `optScopeSelection`.
- Put controls in the array in the order you want them stacked: mode options →
  frame → sub-options → preview toggle → run/help → scope.

---

## 6. The safety chain (do **not** skip or reorder steps)

Every `cmdRun_Click` follows the same sequence. This is what makes the suite
"disaster-proof"; copy it exactly and slot your logic into the marked spot.

```vba
Private Sub cmdRun_Click()
    ' 1. Refuse to run on a protected document
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    ' 2. Shared pre-run guard (auto-save / confirmation). False = user aborted.
    If Not GuardBeforeCleanup("Word Counter") Then Unload Me: Exit Sub
    ' 3. Read preview flag + options
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    ' 4. Resolve scope
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    ' (read your chk/opt options here)
    ' 5. "Nothing selected" guard
    'If Not (optionA Or optionB) Then MsgBox "Nothing selected.", vbInformation: Exit Sub
    ' 6. Preview branch: highlight matches (or show a count), then stop
    If previewOnly Then
        ' ... highlight in wdYellow and count, or build a summary string ...
        MsgBox "Preview complete.", vbInformation
        Unload Me: Exit Sub
    End If
    ' 7. Begin the undo-able transaction
    MarkCleanupStart "Word Counter"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Word Counter"
    On Error GoTo RunErr
    ' 8. >>> DO THE WORK HERE <<<
    Dim results As Collection: Set results = New Collection
    results.Add "Something done: " & 0
    ' 9. Report + close out
    ShowCleanupReport "Word Counter", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
    Unload Me
    Exit Sub
RunErr:
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & Err.Number & " - " & Err.Description
    errMsg = errMsg & vbCrLf & vbCrLf & "Some changes may have been made to the document." & vbCrLf & "Would you like to undo all changes made before the error?"
    If MsgBox(errMsg, vbCritical + vbYesNo, "Cleanup Error") = vbYes Then
        On Error Resume Next: Application.Undo: On Error GoTo 0
    End If
End Sub
```

### Shared helpers available to your tool (in `modCleanupHelpers`)

- `GetTargetRange() As Range` — the current selection as a range (for scope).
- `GuardBeforeCleanup(toolName As String) As Boolean` — pre-run guard; returns
  `False` if the user cancels.
- `MarkCleanupStart(toolName)` / `MarkCleanupEnd()` — status-bar/screen bookends.
- `ShowCleanupReport(title, results As Collection)` — the standard results dialog;
  add `String` lines to a `Collection` and pass it.
- `RemoveAllHighlighting(Optional scopeRange As Range)` — clears yellow highlight
  (use it if your preview highlighting needs resetting).
- `ParagraphsInRange(scopeRange As Range) As Collection` — paragraphs in a range.
- `GetAutoSaveSetting()` / `SetAutoSaveSetting(enabled)` — auto-save preference.

---

## 7. Cautions & gotchas — read every one of these

These are real bugs that have bitten this codebase. Each cost real debugging time.

### 7.1 A single-line `If` cannot contain a block construct (the #1 trap)
This is the most dangerous trap because Word's editor often accepts it visually
but it fails to compile, and naive line-counters miss it.

- **Illegal:** `If x Then With obj: .a = 1: End With` → "Expected: end of statement".
- **Illegal:** `If x Then DoThing: End If` → "End If without block If".
- **Illegal:** `Do: With obj: ...: End With: Loop` (block `With` inside an inline `Do`).

A single-line `If` (anything after `Then` on the same line) may contain **only
simple statements** — assignments, sub calls, `Exit Do/For/Sub`. Any block
keyword (`With`/`End With`, `For`/`Next`, `Do`/`Loop`, `Select Case`/`End Select`,
or a second `If`/`End If`) must be written in **full multi-line block form**:

```vba
If x Then
    With obj
        .a = 1
    End With
End If
```

When in doubt, never inline a block. Multi-line block form is always legal.

### 7.2 Option-button groups need an explicit `GroupName`
Controls are created **flat on the form**, not nested inside their frames, so
Word's automatic "option buttons in the same frame are one group" behavior does
**not** apply. Without intervention, *every* OptionButton on the form is one
mutually-exclusive group. In `UserForm_Initialize`, set a distinct `GroupName`
per logical group, **before** setting any `.Value = True`:

```vba
optModeA.GroupName = "Mode"
optModeB.GroupName = "Mode"
optModeA.Value = True
```

### 7.3 Frames are visual only — they are not containers
Because controls are flat, a frame's `.Visible` toggle will **not** hide the
controls that appear "inside" it. Do not gate behavior on frame visibility. Gate
it on the relevant option's `.Value` in code instead.

### 7.4 Never use highlight color to identify *your own* changes for deletion
Reading `HighlightColorIndex = wdYellow` to decide what to delete will also match
the user's **pre-existing** yellow highlights — silent data loss. Instead, record
the positions (`Range.Start`) of items you detect, then delete by recorded
position.

### 7.5 Iterate backwards when deleting
Deleting rows, columns, paragraphs, footnotes, etc. shifts the collection and
positions. Loop from `.Count` down to `1 Step -1`, and when deleting by recorded
position, sort positions **descending** so earlier deletions never shift the ones
still queued.

### 7.6 Document-level tools must disable selection scope
If the operation is inherently whole-document (metadata, styles, headers/footers,
trailing-paragraph trim), disable selection scope in `UserForm_Initialize`:

```vba
fraScopeSelection.Enabled = False
optScopeDocument.Value = True
optScopeDocument.Caption = "Entire document (always)"
optScopeSelection.Enabled = False
```

### 7.7 Pick the right preview affordance
Most tools preview by **highlighting matches in yellow** and counting them. Some
operations can't be meaningfully highlighted (a font change, header/footer
formatting, metadata) — those preview by showing a **count/summary dialog**
instead. Choose whichever actually informs the user.

### 7.8 Word version-gated APIs
Some methods require newer Word (e.g. `ClearCharacterDirectFormatting` /
`ClearParagraphDirectFormatting` are Word 2013+). If you use one, note the minimum
version and wrap risky calls in `On Error Resume Next` where a graceful skip is
acceptable.

### 7.9 Control creation needs a trust setting
`AttemptCreateControls` only works when **File → Options → Trust Center → Trust
access to the VBA project object model** is enabled. If it is off, control
creation is skipped (and logged); the form still installs but opens empty. This is
expected behavior, not a bug — don't "fix" it by removing the guard.

### 7.10 Auto-layout ordering
`LayoutFormControls` positions controls in the **exact order of the
`ControlsForForm` array**, top to bottom, with command buttons paired two per row.
So order the array the way you want the form to read.

---

## 8. Test checklist before you commit

Static checks are not enough — VBA only truly validates at compile time.

- [ ] Run `python assemble.py` — exits 0 and all five validators report clean.
- [ ] **Debug → Compile** the installed project in the VBA editor. Zero errors.
- [ ] Run `InstallCleanupSuite` on a clean document; confirm the new form and its
      launcher button appear.
- [ ] Open the tool, run in **Preview** mode first on a throwaway copy.
- [ ] Run for real, then press **Ctrl+Z** — confirm the change undoes cleanly.
- [ ] Run on a **protected** document — confirm it is blocked with the guard message.
- [ ] Run with **no selection** and with a **selection** — confirm scope behaves.
- [ ] Run with **nothing selected** in the options — confirm the "nothing selected"
      guard fires instead of doing work.
- [ ] Confirm the **Help** button shows accurate text.
- [ ] Add the new form name to `IsSuiteComponent`'s list (§3, step 8).
- [ ] Re-run over an existing install via `UninstallCleanupSuite` then `InstallCleanupSuite` — see §9.

---

## 9. Known caveats (not bugs to "fix" casually)

- **Re-running the installer** over an existing install hits VBA's deferred
  component-removal quirk (duplicate `frmX1` components / intermittent failure),
  because removing and re-adding a component in one pass is inherently racy in
  VBIDE. **Use the developer tools:** run `UninstallCleanupSuite` first (it only
  removes, so the deletions commit when it returns), then run `InstallCleanupSuite`
  as a separate run. `ReinstallCleanupSuite` does both in one click via
  `Application.OnTime` (needs the host code in a standard module). A first run on a
  clean project needs neither.
- **Auto-layout is functional, not hand-tuned.** Forms come out as a clean single
  column. If you want pixel-perfect layout, adjust in the designer after install.
- **The pasted module is what end users see.** Keep end-user-facing comments
  minimal; put developer documentation here, not in the `.txt`.

---

## 10. Points of interest / design rationale

- **Why string-builder functions instead of `.frm` files?** So the whole suite is
  one copy-pasteable block requiring no file import — the lowest-friction install
  for non-technical users.
- **Why a custom `UndoRecord` per tool?** It groups each tool's edits into a single
  named undo step, so one Ctrl+Z reverts the whole operation, and the error handler
  can offer a clean rollback.
- **Why preview-by-highlight?** It lets users see exactly what *would* change before
  committing, using Word's own highlight as a zero-risk visual diff.
- **Why the protection + guard + nothing-selected gates?** Defense in depth: the
  tool refuses unsafe states early rather than producing surprising edits.

When in doubt, copy the closest existing tool: a **scoped** tool (e.g. the Spacing
or Table Cleaner forms) or a **document-level** tool (e.g. Metadata Scrubber or the
Footnote/Endnote Remover), whichever matches your operation.
