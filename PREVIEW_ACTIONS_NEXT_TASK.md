# Preview Actions Duplicate Text — Next Task for Codex

This is the next known UI issue to investigate after the safety setup and repo/local sync check are complete.

## Important prerequisite

First read:

```text
REPO_FINDINGS_AND_SYNC_CHECK.md
```

GitHub `main` may be behind Chris's local folder or the actual `.docm`. The screenshot shows a three-button Preview Actions bar, while GitHub `main` appears to show an older two-button source design. Do not fix the wrong version.

## Symptom

The Preview Actions dialog/action bar is still showing an unwanted extra `Preview Actions` text inside the form body/client area.

The screenshot showed two visible instances:

1. The native UserForm/window caption at the very top. This comes from `frmPreviewActions.Caption`, the form `Caption` property, or installer styling.
2. A separate client-area caption/control just below the title bar, at the upper-left of the form. This is the unwanted duplicate.

Earlier conversation wording said there had been three instances. The latest screenshot showed the duplicate problem still remains as two visible instances: title bar plus in-body text.

## Critical observation

Do not keep targeting only `lblTitle`.

Previous audits reported:

```text
lblTitle: 0,0,0,0, hidden, blank
```

If that audit is accurate at runtime, then the remaining client-area `Preview Actions` text is almost certainly not `lblTitle`.

## GitHub-main clue

In GitHub `main`, the installer maps `frmPreviewActions` to a friendly form caption of `Preview Actions`, and the installer also sets the form/designer caption. That explains the native title-bar text, but it does not by itself prove what is drawing the extra client-area text in Chris's screenshot.

The screenshot's three-button UI also does not match the current GitHub-main `frmPreviewActions` shape. That strongly suggests local/uncommitted/unpushed or embedded `.docm` state must be inspected before editing.

## Likely causes

The remaining in-body/client-area text may be one of these:

- a `Frame.Caption`
- a `Page.Caption` or `MultiPage` caption
- another label with a different name
- a dynamically created control at runtime
- a stale embedded copy of the form inside the `.docm` that is not in sync with exported source files
- initialization code that restores `.Caption = "Preview Actions"` after the static form layout loads
- installer code rebuilding or reimporting the form from stale layout data
- tests or generated/distributable code that were not updated with the local UI changes

## What not to do

- Do not assume the visible text is `lblTitle`.
- Do not make another layout-only change before proving which runtime control prints the visible text.
- Do not rely only on GitHub `main`, because it may be stale relative to Chris's local project.
- Do not rely only on exported `.bas` source files if the deliverable `.docm` has embedded forms.
- Do not report success until the actual working `.docm` and/or runtime form behavior has been checked.

## Runtime audit to add temporarily

Add this in `frmPreviewActions.Initialize`, or immediately before `.Show`, and capture the Immediate Window output.

```vb
Dim ctl As MSForms.Control
Debug.Print "FORM CAPTION:", Me.Caption

For Each ctl In Me.Controls
    On Error Resume Next
    Debug.Print ctl.Name, TypeName(ctl), ctl.Caption, ctl.Visible, ctl.Left, ctl.Top, ctl.Width, ctl.Height
    On Error GoTo 0
Next ctl
```

If nested controls are possible, also recursively inspect controls inside frames/multipages/pages.

## Project-wide searches to run

Search the entire project, including source files, installer code, tests, generated distributable text, and any scripts that rebuild or import forms:

```text
Preview Actions
.Caption
Controls.Add
Frame
MultiPage
Page
lblTitle
frmPreviewActions
cmdPreview
cmdReconfigure
cmdApply
cmdClear
```

Also inspect the embedded `.docm` VBA/forms if the project uses generated/exported source files plus a working Word document.

## Fix strategy

1. Create a backup zip first.
2. Read `AI_HANDOFF.md`, `AI_RULES.md`, and `REPO_FINDINGS_AND_SYNC_CHECK.md` first.
3. Determine whether GitHub `main`, local source, generated `VBA_Cleanup_tool.txt`, or embedded `.docm` is current truth.
4. Run the runtime control audit.
5. Identify the actual runtime control or form property responsible for the in-body `Preview Actions` text.
6. Fix that exact source.
7. Confirm source files, generated distributable, tests, and working `.docm` are in sync.
8. Re-run validators/tests.
9. Manually open the Preview Actions dialog on Chris's high-DPI environment if possible.
10. Append a detailed handoff entry explaining exactly which control/property caused the duplicate.

## Acceptance criteria

- The window title may still say `Preview Actions` unless Chris asks to blank the native form caption.
- Inside the form body, the upper-left duplicate `Preview Actions` text should be gone.
- The intended tool name, such as `Capitalization Fixer`, should remain visible.
- Buttons should remain visible and not clipped at 200% Windows scaling.
- The summary/status line should remain visible and not clipped.
- The fix should be verified against the actual `.docm` used by Chris, not only against exported source.
- `AI_HANDOFF.md` should say whether GitHub `main`, local files, generated text, and `.docm` were in sync or not.
