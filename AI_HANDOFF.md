# AI Handoff Log — CleanupSuite for MS Word

Project folder:

```text
C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word
```

GitHub repo:

```text
https://github.com/MasseysLab/CleanupSuiteForWord
MasseysLab/CleanupSuiteForWord
Default branch observed by ChatGPT: main
```

This file is the first thing Codex or ChatGPT should read before editing the project.

---

## Session: 2026-06-12 04:22:14
Agent: ChatGPT

Goal:
- Create and revise a safety package so Chris can alternate between Codex and ChatGPT without the project being harmed.
- Make the verification/double-check requirements explicit.
- Include the Preview Actions duplicate-text investigation as the next known task for Codex.
- Add a v3 repo-specific warning after inspecting GitHub: GitHub `main` may not match Chris's current local working folder or actual `.docm`.

Backup created:
- Not created by ChatGPT because ChatGPT cannot access Chris's local Windows project folder directly.
- Codex should create the first real backup zip at the start of its next local session before copying/editing anything.

Files intended to be added to the project root:
- `AI_HANDOFF.md`
- `AI_RULES.md`
- `CODEX_SETUP_TASK.md`
- `GITHUB_WORKFLOW.md`
- `REPO_FINDINGS_AND_SYNC_CHECK.md`
- `PREVIEW_ACTIONS_NEXT_TASK.md`
- `Start-AI-Session.ps1`
- Optional/reference: `README_FOR_CHRIS.txt`

## Critical repo/local sync warning

ChatGPT inspected GitHub repo `MasseysLab/CleanupSuiteForWord` on branch `main`. GitHub `main` appears to show an older Preview Actions design than Chris's screenshot and the latest local Codex conversation.

Chris's screenshot shows three Preview Actions buttons:

```text
Preview is ON    Reconfigure    Apply
```

GitHub `main` appears to still have a two-button `frmPreviewActions` source design with controls like:

```text
cmdApply
cmdClear
```

Therefore Codex must not assume GitHub `main` is the newest truth. Before fixing Preview Actions, compare:

1. GitHub `main`.
2. Chris's local working tree.
3. Generated `VBA_Cleanup_tool.txt`.
4. The actual embedded or installed `.docm`/VBA project Chris is running.

See `REPO_FINDINGS_AND_SYNC_CHECK.md` before editing the Preview Actions bug.

## Important message for Codex

Codex, please follow this project safety protocol from now on.

At the start of every session, before changing anything:

1. Read `AI_HANDOFF.md` if it exists.
2. Read `AI_RULES.md` if it exists.
3. Read `REPO_FINDINGS_AND_SYNC_CHECK.md` if working on Preview Actions or repo synchronization.
4. Create a timestamped backup zip of the full project folder.
5. Do not overwrite previous backups.
6. Check Git remote, current branch, and status.
7. Then inspect the current project state before editing.

Backup naming format:

```text
backups\CleanupSuite_backup_YYYY-MM-DD_HH-mm-ss.zip
```

## Required double-checks before Codex reports setup complete

Codex should not just copy files and say done. It should verify that the safety setup actually works:

- Confirm every required file exists in the project root.
- Confirm `AI_HANDOFF.md`, `AI_RULES.md`, and `REPO_FINDINGS_AND_SYNC_CHECK.md` are readable.
- Confirm a backup zip was created in `backups\`.
- Confirm the new backup zip does not include the `backups` folder inside itself.
- Confirm `backups/` is ignored by Git or otherwise will not be committed.
- Confirm Git remote URL, current branch, and working tree status.
- Append a Codex setup entry to `AI_HANDOFF.md` that records what was done.
- If anything cannot be verified, say exactly what could not be verified.

## After every session where changes are made

Append a detailed note to `AI_HANDOFF.md`.

Each handoff entry should include:

```markdown
## Session: YYYY-MM-DD HH:mm
Agent: Codex / ChatGPT

Goal:
- What the user asked for.

Backup created:
- backups\CleanupSuite_backup_YYYY-MM-DD_HH-mm-ss.zip

Git/repo state:
- Remote:
- Branch:
- Status before:
- Status after:
- GitHub/local/.docm sync checked: yes/no

Files changed:
- pathile1.bas
- pathile2.frm
- pathile3.py

Summary of changes:
- Explain exactly what was changed and why.

Important observations:
- Mention confusing behavior, suspected causes, scaling issues, Word/VBA quirks, or things intentionally left alone.

Tests/checks run:
- VBA validator result
- Python tests result
- Manual inspection result
- `.docm` updated: yes/no
- Git status/diff checked: yes/no

Known remaining issues:
- Anything still broken.
- Anything not attempted.
- Anything the next agent should be careful about.

Next recommended step:
- The safest next action.
```

## Specific next task: Preview Actions duplicate text

Read both files first:

```text
REPO_FINDINGS_AND_SYNC_CHECK.md
PREVIEW_ACTIONS_NEXT_TASK.md
```

Summary:
- The screenshot showed a native title-bar `Preview Actions` and an unwanted client-area `Preview Actions` duplicate.
- Previous audits claimed `lblTitle` was hidden, blank, and sized `0,0,0,0`.
- Therefore the remaining client-area `Preview Actions` text is probably not `lblTitle`.
- GitHub `main` may be behind Chris's local state; the screenshot's three-button UI does not match GitHub-main's two-button-looking source.
- Do not keep changing `lblTitle` unless runtime audit proves it is the visible control.
- Run a runtime controls audit and compare GitHub/local/generated `.txt`/`.docm` before changing more layout.

Runtime audit snippet:

```vb
Dim ctl As MSForms.Control
Debug.Print "FORM CAPTION:", Me.Caption

For Each ctl In Me.Controls
    On Error Resume Next
    Debug.Print ctl.Name, TypeName(ctl), ctl.Caption, ctl.Visible, ctl.Left, ctl.Top, ctl.Width, ctl.Height
    On Error GoTo 0
Next ctl
```

Known remaining issues:
- ChatGPT has not directly edited Chris's local project folder.
- Codex should create the first real backup zip before making further changes.
- The Preview Actions duplicate issue remains the next recommended technical investigation.
- Repo/local/generated/.docm sync must be checked before fixing Preview Actions.

Next recommended step:
- Copy the files from this package into the project root without destroying any newer existing handoff files.
- Run `Start-AI-Session.ps1` or otherwise create a backup zip.
- Perform the required double-checks above.
- Then continue with `REPO_FINDINGS_AND_SYNC_CHECK.md` and `PREVIEW_ACTIONS_NEXT_TASK.md`.

## Session: 2026-06-11 22:29
Agent: Codex

Goal:
- Install the v3 AI collaboration safety files from `codex_cleanup_suite_prep_package_v3.zip`.
- Create and verify the first local backup zip for this workflow.
- Record Git/repo state and setup verification so Codex and ChatGPT can alternate more safely.

Backup created:
- `backups\CleanupSuite_backup_2026-06-11_22-29-35.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before setup: existing dirty working tree from current UI/VBA work; only `.gitignore` existed from the new AI setup file list.
- Status after setup: existing dirty working tree preserved; `.gitignore` modified; AI safety files added as untracked files; backup zip ignored by Git.
- GitHub/local/generated/.docm sync checked: partially. The package warning was read and recorded; a full Preview Actions runtime sync investigation was not performed as part of setup.

Files changed:
- `.gitignore`
- `AI_HANDOFF.md`
- `AI_RULES.md`
- `CODEX_SETUP_TASK.md`
- `GITHUB_WORKFLOW.md`
- `REPO_FINDINGS_AND_SYNC_CHECK.md`
- `PREVIEW_ACTIONS_NEXT_TASK.md`
- `Start-AI-Session.ps1`
- `README_FOR_CHRIS.txt`

Summary of changes:
- Copied the v3 safety/handoff files from the prep package into the project root.
- Added `backups/` to `.gitignore`.
- Created a timestamped backup zip before modifying the project root.
- Preserved the existing dirty working tree and did not reset, pull, clean, or overwrite code changes.

Important observations:
- No prior `AI_HANDOFF.md`, `AI_RULES.md`, or related safety files existed in the project root, so there was no prior handoff history to merge.
- The backup zip was created after current project work was already dirty, so it preserves the current local in-progress state.
- GitHub `main` should not be treated as the current truth for Preview Actions or `.docm` behavior without comparing against local/generated/embedded state.

Tests/checks run:
- Confirmed every required safety file exists in the project root.
- Confirmed `AI_HANDOFF.md`, `AI_RULES.md`, `CODEX_SETUP_TASK.md`, `GITHUB_WORKFLOW.md`, `REPO_FINDINGS_AND_SYNC_CHECK.md`, `PREVIEW_ACTIONS_NEXT_TASK.md`, `Start-AI-Session.ps1`, and `README_FOR_CHRIS.txt` are readable.
- Confirmed backup zip exists.
- Confirmed backup zip does not include the `backups/` folder.
- Confirmed `backups/` is ignored by Git.
- Confirmed Git remote, branch, recent log, and status.

Known remaining issues:
- The project has many pre-existing modified files from the current CleanupSuite UI/VBA work.
- Preview Actions should still be manually/runtime verified on Chris's high-DPI Word environment after any further changes.
- No Git commit or push was made during this setup.

Next recommended step:
- Continue regular CleanupSuite work only after reading `AI_RULES.md`; for Preview Actions or sync issues, read `REPO_FINDINGS_AND_SYNC_CHECK.md` and `PREVIEW_ACTIONS_NEXT_TASK.md` first.

## Session: 2026-06-11 22:42
Agent: Codex

Goal:
- Create a fresh handoff so Chris can safely start a new Codex chat after seeing context condensation.
- Record the current co-op model with ChatGPT: Chat can review GitHub and write planning notes, but Codex/local must apply and verify changes against the real Windows project and `.docm`.

Backup created:
- `backups\CleanupSuite_backup_2026-06-11_22-41-57.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree with existing CleanupSuite UI/VBA changes plus newly added AI safety files.
- Status after: same dirty working tree; only this handoff entry was added during this mini-session.
- GitHub/local/generated/.docm sync checked: no. This was a handoff/documentation update only.

Files changed:
- `AI_HANDOFF.md`

Summary of changes:
- Appended this fresh-session handoff entry.
- Confirmed the best collaboration split: ChatGPT should use GitHub/docs for review, planning, and proposals; Codex should be the one to edit local VBA, generated files, embedded `.docm`, backups, tests, commits, and pushes.
- Recorded that a prescribed GitHub folder for Chat notes is acceptable, but should be treated as advisory unless Codex verifies against local source/generated/embedded project state.

Important observations:
- Do not assume GitHub `main` is identical to Chris's current local folder or to the `.docm` Word is actually running.
- The current project has many pre-existing modified files from the ongoing launcher/action-bar/UI work. Do not reset or overwrite them.
- Latest known Preview Actions intent: compact action bar with native caption `Preview Actions`, centered tool name, buttons `Preview is ON/OFF`, `Reconfigure`, and `Apply`, plus a summary line that must remain visible at high DPI.
- Latest known technical concern: Chris's high-DPI machine is 3000 x 2000 with Windows scaling at 200%; form sizing must be verified visually/runtime, not only by source inspection.
- Latest known functional concern: applying on a blank document previously crashed Word; local code has a blank-document guard in `frmPreviewActions`/`modCleanupHelpers`, but Chris should still runtime-test it.

Tests/checks run:
- Read `AI_HANDOFF.md`.
- Read `AI_RULES.md`.
- Created backup zip before editing.
- Checked Git remote, branch, and status.
- No VBA validator, Python tests, `.docm` sync audit, or Word runtime test was run in this mini-session.

Known remaining issues:
- The uncommitted UI/VBA work still needs final runtime verification in Word.
- The action bar may still need visual adjustment if Chris's latest runtime screenshot does not match the intended compact layout.
- ChatGPT cannot see the local Windows folder, so Chat-authored advice must be applied cautiously.

Next recommended step:
- In a new Codex session, start by reading `AI_HANDOFF.md`, `AI_RULES.md`, `REPO_FINDINGS_AND_SYNC_CHECK.md`, and `PREVIEW_ACTIONS_NEXT_TASK.md`; then create a fresh backup, inspect current Word/runtime behavior, and continue the action-bar verification before committing or pushing.

## Session: 2026-06-11 22:50
Agent: Codex

Goal:
- Add explicit startup permission guidance so a fresh Codex session asks Chris for needed permissions up front instead of interrupting him at odd moments.

Backup created:
- `backups\CleanupSuite_backup_2026-06-11_22-49-36.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree with existing CleanupSuite UI/VBA changes plus AI safety files.
- Status after: same dirty working tree; only `AI_HANDOFF.md` was updated during this mini-session.
- GitHub/local/generated/.docm sync checked: no. This was a handoff/documentation update only.

Files changed:
- `AI_HANDOFF.md`

Summary of changes:
- Added this startup permission instruction for the next Codex session.

Startup permission request for next Codex session:
- At the beginning of the next session, after reading `AI_HANDOFF.md` and `AI_RULES.md`, ask Chris for the following permissions in one clear prompt:

```text
For this CleanupSuite session, may I have permission to:

1. Read and edit files in:
   C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word

2. Run PowerShell/shell commands needed for backups, tests, validators, file inspection, VBA export/import helpers, and Git status/diff commands.

3. Create timestamped backup zip files in:
   C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word\backups

4. Update the generated VBA bundle and working/practice `.docm` files when needed for the task.

5. Use network/GitHub access for fetch/compare/push only when the task calls for it.

6. Proceed without asking for every small safe edit inside the CleanupSuite project folder.

I will still avoid destructive operations unless you explicitly ask for them or I explain first. I will not reset, clean, delete major files, overwrite the `.docm`, pull, merge, rebase, or push unless the task specifically calls for it or you confirm it.
```

Recommended default if Chris says yes:
- Treat the session as approved for normal CleanupSuite development work in the local project folder.
- Still create a fresh backup before edits.
- Still preserve uncommitted work.
- Still ask before destructive Git/file operations.
- Still ask or clearly announce before pushing to GitHub unless Chris's current task explicitly says to push.

Important observations:
- Current session already had full local access, but the next session may start with different permissions or a different approval prompt.
- Chris wants permission friction handled at startup, not during the middle of focused UI/VBA work.

Tests/checks run:
- Created backup zip before editing.
- Updated `AI_HANDOFF.md`.
- Did not run VBA validator, Python tests, `.docm` sync audit, or Word runtime test because this was documentation-only.

Known remaining issues:
- No code behavior changed.
- The uncommitted CleanupSuite UI/VBA work still needs normal verification before commit/push.

Next recommended step:
- Start the next Codex session by reading this handoff, asking the startup permission prompt above, creating a fresh backup after permission is granted, and then continuing the Preview Actions/action-bar verification work.

## Session: 2026-06-11 23:56
Agent: Codex

Goal:
- Find and fix the Preview Actions action-bar duplicate in-body `Preview Actions` text.

Backup created:
- `backups\CleanupSuite_backup_2026-06-11_23-31-23.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree with existing CleanupSuite UI/VBA changes plus AI safety files.
- Status after: same dirty working tree; this session changed the Preview Actions installer/test/generated bundle path and updated both workspace `.docm` files.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`; GitHub/main remains older than local working state.

Files changed:
- `src\installer\installer.bas`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Root cause found: the duplicate was not a visible `lblTitle` control after `Configure`; it came from the persisted design-time MSForms UserForm caption for `frmPreviewActions`.
- A visual `PrintWindow` capture reproduced the bug even though `lblTitle` was blank/hidden. Temporarily blanking runtime `Me.Caption` removed only the native title-bar text, not the in-body duplicate. Temporarily blanking `frmPreviewActions.Designer.Caption` removed the in-body duplicate while keeping the runtime/native title bar.
- Updated `LayoutPreviewActionsControls` so `frmPreviewActions` keeps its design-time caption blank (`comp.Properties("Caption") = ""` and `designer.Caption = ""`). Runtime code still sets `Me.Caption = "Preview Actions"`.
- Added a focused regression assertion in `tests\test_preview_action_panel.py`.
- Rebuilt `VBA_Cleanup_tool.txt` with `assemble.py` and synced the practice generated bundle.
- Updated both embedded `.docm` copies so `frmPreviewActions.Designer.Caption` is blank.

Important observations:
- Old Git baseline/GitHub-era code explicitly set `frmPreviewActions.lblTitle` to `Preview Actions` and positioned it visibly; current source had already fixed `lblTitle`, but the embedded design-time form caption still produced a visible in-body caption.
- Window-only visual capture after the fix showed the in-body duplicate gone.
- The action-bar title area may still need a high-DPI layout pass: the capture suggests the centered tool name can appear far to the right/clipped under the automation capture, although the duplicate caption itself is fixed.
- Temporary audit modules and screenshot artifacts were removed; final verification found no leftover `modCodex*` modules in either `.docm`.

Tests/checks run:
- Created backup zip before edits.
- Runtime controls audit before fix on both `.docm` files.
- Window-only visual captures before and after fix.
- Verified both `.docm` files have `frmPreviewActions.Designer.Caption=[]`, 6 controls, no `modCodex*` temporary modules, and saved state true.
- Ran `python assemble.py`; validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Ran bundled Python focused test: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Ran bundled Python full test discovery: `python -m unittest discover -s tests` -> 14 tests OK.
- Checked no `WINWORD` process remained after automation cleanup.
- Checked Git status/diff context.

Known remaining issues:
- The project still has many pre-existing modified files unrelated to this specific duplicate-caption fix.
- A final human-visible high-DPI Word check is still recommended for spacing/clipping, especially around the tool-name label and summary line.

Next recommended step:
- Open Preview Actions normally in Word on Chris's 3000 x 2000 / 200% scaling display and confirm the duplicate caption is gone; if the tool name or buttons are clipped, do a separate action-bar sizing/layout pass.

## Session: 2026-06-12 00:14
Agent: Codex

Goal:
- Fix the Preview Actions action-bar sizing/clipping found during visual verification.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_00-08-10.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree with existing CleanupSuite UI/VBA changes, AI safety files, and the previous Preview Actions duplicate-caption fix.
- Status after: same dirty working tree; this session changed the Preview Actions source/test/generated bundle path and updated both workspace `.docm` files.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `src\forms\frmPreviewActions.bas`
- `src\installer\installer.bas`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Visual verification after the duplicate-caption fix showed the in-body `Preview Actions` text was gone, but the tool name and action buttons were clipped because the outer form shell was too small under the 200% scaling render path.
- Added a failing regression expectation first, then widened the Preview Actions outer UserForm shell to `FORM_W = 820` and `FORM_H = 260` while preserving the existing compact internal coordinate layout via `CONTENT_W = 430`.
- Updated both source form code and installer-generated design layout.
- Rebuilt `VBA_Cleanup_tool.txt`, synced the practice generated bundle, and updated both embedded `.docm` copies with the new form code and dimensions.

Important observations:
- The successful visual capture after the sizing change showed: native `Preview Actions` title only, centered `Capitalization Fixer`, visible `Preview is ON`, `Reconfigure`, and `Apply` buttons, and visible summary text.
- The window-only `PrintWindow` capture still causes hidden Word automation to sometimes time out during cleanup; orphaned `/Automation -Embedding` Word processes were explicitly inspected and closed.
- Temporary screenshots, stale Word lock files, and temporary audit modules were removed.

Tests/checks run:
- Created backup zip before edits.
- TDD red check: focused preview action test failed before production changes because expected `FORM_W = 820`, `FORM_H = 260`, and `CONTENT_W = 430` were absent.
- Focused test after code changes: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Visual window-only capture confirmed the resized action bar fits.
- Full tests: `python -m unittest discover -s tests` -> 14 tests OK.
- Verified both `.docm` files read-only: `DesignerCaption=[]`, width about `828.1`, height about `260.05`, 6 controls, no `modCodex*` temporary modules, saved state true.
- Confirmed no `WINWORD` process remained.

Known remaining issues:
- The project still has many pre-existing modified files unrelated to this focused action-bar sizing fix.
- Final hands-on Word testing is still useful before release, but Codex visually verified the action bar under the current automation/display path.

Next recommended step:
- Continue toward final verification/commit prep for the broader uncommitted CleanupSuite UI/VBA work, or run through normal tool workflows in the practice `.docm` to check end-to-end preview/apply behavior.

## Session: 2026-06-12 00:46
Agent: Codex

Goal:
- Turn the Preview Actions oversized action box into a compact preview action bar with tight button/text spacing and aligned centers.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_00-36-35.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree with existing CleanupSuite UI/VBA changes and prior Preview Actions fixes.
- Status after: source/tests/generated text bundles and both embedded `.docm` copies updated for the compact bar.

Files changed:
- `src\forms\frmPreviewActions.bas`
- `src\installer\installer.bas`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `AI_HANDOFF.md`

Summary of changes:
- Replaced the oversized Preview Actions layout target with a compact bar target: runtime form `FORM_W = 530`, `FORM_H = 160`, `CONTENT_W = 250`, margin `8`, button width `80`, button height `17`, gap `5`.
- Tightened vertical layout: tool name at top `2`, buttons at `17`, summary at `39`.
- Kept the center-line math aligned: the tool-name label width equals the 3-button group width, and the middle `Reconfigure` button center is the same as the tool-name center.
- Reduced visual weight: tool name font from 11 to 9, summary from 10 bold to 8.5 regular, buttons from 10 to 8.5.
- Updated installer-generated layout constants, rebuilt/synced generated text bundles, and updated both embedded `.docm` copies.

Important observations:
- A runtime visual capture of the intermediate 580/180 layout showed correct centering and no clipping, but it still read as a chunky mini-dialog rather than a compact bar.
- A runtime visual capture of the 460/135 layout showed the desired compact feel but clipped the `Apply` button and summary under Word's high-DPI scaling.
- The final 530/160 shell preserves the same tight internal control layout while giving Word enough outer shell to avoid clipping.
- Final visual capture showed centered `Capitalization Fixer`, aligned `Reconfigure`, visible `Preview is ON`, `Reconfigure`, `Apply`, and visible summary text.

Tests/checks run:
- Focused TDD red check existed before production changes for the compact layout.
- Focused test after source changes: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Updated both `.docm` copies with the final embedded `frmPreviewActions` code/control positions.
- Visual window-only capture after final `.docm` update: `757x229` pixels, no right-side or bottom clipping.
- Full tests: `python -m unittest discover -s tests` -> 14 tests OK.
- Verified both `.docm` files read-only: final width/height constants present, `DesignerCaption=[]`, 6 controls, no `Codex*` or `modCodex*` temporary modules, saved state true.
- Confirmed no `WINWORD` process remained.
- Removed temporary visual-capture screenshot.

Known remaining issues:
- The project still has many pre-existing modified files unrelated to this focused action-bar sizing fix.
- Final hands-on Word testing is still useful before release, but Codex visually verified the compact bar under the current automation/display path.

Next recommended step:
- Run through a normal preview/reconfigure/apply workflow in the practice `.docm` to confirm the action bar feels right during actual editing.
