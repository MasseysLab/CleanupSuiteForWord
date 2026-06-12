# Repo Findings and Sync Check — CleanupSuiteForWord

This file was added after ChatGPT inspected the GitHub repo Chris provided:

```text
https://github.com/MasseysLab/CleanupSuiteForWord
MasseysLab/CleanupSuiteForWord
Default branch observed: main
```

## Critical finding

GitHub `main` may not match Chris's current local working folder or the actual working `.docm`/VBA project.

The screenshot Chris showed has a Preview Actions UI with three buttons:

```text
Preview is ON    Reconfigure    Apply
```

But the GitHub `main` source that ChatGPT inspected still appears to have an older/two-button `frmPreviewActions` design using controls such as:

```text
cmdApply
cmdClear
```

and tests that expect:

```text
ControlsForForm = Array("lblTitle", "lblSummary", "lblHint", "cmdApply", "cmdClear")
```

Therefore: do not assume GitHub `main` is the newest truth for the Preview Actions bug. Local files or the embedded `.docm` may contain newer uncommitted/unpushed changes.

## Do not destroy local changes

Before running any command that can change the working tree, Codex should:

1. Create a timestamped backup zip.
2. Record Git branch and status.
3. Record whether local files differ from GitHub/main.
4. Avoid `git reset --hard`, `git clean`, checkout/switch, pull, merge, or rebase until Chris's local state is backed up and documented.

## Required sync check before fixing Preview Actions

Run these in the local project folder and record the output summary in `AI_HANDOFF.md`:

```powershell
git remote -v
git branch --show-current
git status --short
git log --oneline -5
git diff -- src/forms/frmPreviewActions.bas src/installer/installer.bas tests/test_preview_action_panel.py VBA_Cleanup_tool.txt
```

Then search the local working tree:

```powershell
git grep -n -e "Preview Actions" -e "frmPreviewActions" -e "lblTitle" -e "cmdPreview" -e "cmdReconfigure" -e "cmdApply" -e "cmdClear"
```

If `git grep` misses generated or binary/embedded content, also use PowerShell search over text files:

```powershell
Get-ChildItem -Recurse -File | Select-String -Pattern "Preview Actions","frmPreviewActions","lblTitle","cmdPreview","cmdReconfigure","cmdApply","cmdClear"
```

## Files to compare first

At minimum, compare these local files against GitHub/main and against the actual installed `.docm` behavior:

```text
src/forms/frmPreviewActions.bas
src/installer/installer.bas
tests/test_preview_action_panel.py
src/manifest.txt
VBA_Cleanup_tool.txt
```

## Embedded .docm warning

This project builds from text source into a distributable `.txt`, and the actual Word/VBA runtime may also involve an embedded `.docm` or currently installed VBA project. The visible bug is runtime UI behavior, so source-only checks are not enough.

Codex should verify whether the actual `.docm` or installed VBA project is in sync with:

```text
src/forms/frmPreviewActions.bas
src/installer/installer.bas
VBA_Cleanup_tool.txt
```

## Updated instruction for Codex

Before solving the Preview Actions duplicate problem, determine which of these is current truth:

1. GitHub `main`.
2. Chris's local working tree.
3. Generated `VBA_Cleanup_tool.txt`.
4. The embedded or installed `.docm`/VBA project Chris is actually running.

If they differ, do not guess. Record the differences in `AI_HANDOFF.md`, preserve the newest useful work, and ask Chris before discarding anything.
