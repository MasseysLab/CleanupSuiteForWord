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

## Session: 2026-06-14 12:45
Agent: Codex

Goal:
- Add ChatGPT's `cleanup_suite_docs_0_7_4_addon.zip` documentation package to the locked 0.7.4 release.
- Keep the work documentation-only and avoid code behavior changes.
- Compare package wording against the actual local 0.7.4 UI/tool behavior before committing.

Backup created:
- `backups\CleanupSuite_backup_2026-06-14_12-39-54_before-docs-0.7.4-addon.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: clean working tree.
- Status after: documentation files and `AI_HANDOFF.md` changed, pending commit at the time this entry was written.
- GitHub/local/generated/.docm sync checked: local source/tests were checked for the documented UI labels; no GitHub sync or `.docm` rewrite was needed because this was docs-only.

Files changed:
- `docs\Documentation_Addon_Handoff_v0.7.4.md`
- `docs\Human_Friendly_User_Manual_v0.7.4.md`
- `docs\Programmers_Guide_Adding_Tools_v0.7.4.md`
- `docs\Tool_Reference_Refined_v0.7.4.md`
- `documents\CleanupSuite_Human_Friendly_User_Manual_v0.7.4.docx`
- `documents\CleanupSuite_Programmers_Guide_Adding_Tools_v0.7.4.docx`
- `AI_HANDOFF.md`

Summary of changes:
- Added the Markdown documentation package under `docs\`.
- Added the DOCX documentation package under `documents\`.
- Did not add screenshots.
- Did not copy package wrapper files `CODEX_ADD_0.7.4_DOCS_TASK.md` or `README_DOC_PACKAGE.md` into the repo because the package task listed the release docs themselves as the files to add.
- Removed package Markdown trailing whitespace so `git diff --check` passes; this was formatting-only.

Documentation wording adjusted after local 0.7.4 comparison:
- In the user manual Markdown, changed the older setting label `Return to main menu after apply` to the current 0.7.4 pair:
  - `Show completion review after apply`
  - `Return to main menu after completion review`
- In Markdown and the user-manual DOCX, changed `Direct Formatting Stripper` section heading to `Formatting Stripper / Strip Formatting` to reflect both the current form name and launcher label.
- In Markdown and the user-manual DOCX, changed `Remove Objects & Elements` section heading to `Object Remover / Objects` to reflect the current form name and launcher label.

Important observations:
- Local 0.7.4 source/tests confirm current labels for `Preview is ON`, `Preview is OFF`, `Reconfigure`, `Apply`, `Show completion review after apply`, and `Return to main menu after completion review`.
- `src\forms\frmFormattingStripper.bas` still contains older help text that says `Direct Formatting Stripper`, while shared/current guided UI naming uses `Formatting Stripper`; the docs use a combined heading to avoid misleading either route.
- This change intentionally does not touch source code, generated VBA bundles, or `.docm` files.

Tests/checks run:
- Confirmed the package zip exists at `C:\Users\Chris\Downloads\cleanup_suite_docs_0_7_4_addon.zip`.
- Created and verified a backup zip; confirmed it does not contain the `backups` folder.
- Extracted the package to a temp folder before copying.
- Compared docs against local source/tests with `rg` for 0.7.4 labels and tool names.
- Verified copied Markdown contains the adjusted 0.7.4 wording.
- Verified both DOCX files are readable with bundled Python `python-docx`.
- Attempted DOCX render QA with bundled `render_docx.py`, but rendering failed because the renderer could not find its LibreOffice/soffice executable. No PNG visual QA was available from that tool.

Known remaining issues:
- DOCX visual render QA was not completed because the renderer dependency was unavailable.
- Screenshots are intentionally deferred for a later documentation pass.

Next recommended step:
- Run the normal test suite and git whitespace/status checks, then commit only these documentation and handoff changes with `Add 0.7.4 user and developer documentation`.

## Session: 2026-06-14 14:51
Agent: Codex

Goal:
- Correct the release plan so `0.7.4` remains the programming-side/code-freeze checkpoint and `0.7.5` becomes the official documentation-complete release.
- Keep feature/tool behavior identical to `0.7.4`.
- Update documentation, version wording, `SUITE_VERSION`, generated bundles, and embedded `.docm` version metadata so Word identifies the documented release as `0.7.5`.
- Commit and push the `0.7.5` documentation-release preparation.

Backup created:
- `backups\CleanupSuite_backup_2026-06-14_14-44-24_before-v0.7.5-doc-release.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: clean branch at pushed commit `c48038f Add 0.7.4 user and developer documentation`.
- Status after: pending commit at the time this entry was written.
- GitHub/local/generated/.docm sync checked: yes for version metadata. Source, generated root bundle, practice bundle, and both embedded `.docm` files were checked for `SUITE_VERSION = "0.7.5"`.

Files changed:
- `README.md`
- `VERSIONING.md`
- `src\modules\modCleanupLauncher.bas`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `tests\test_launcher_redesign.py`
- `docs\Documentation_Addon_Handoff_v0.7.5.md`
- `docs\Human_Friendly_User_Manual_v0.7.5.md`
- `docs\Programmers_Guide_Adding_Tools_v0.7.5.md`
- `docs\Tool_Reference_Refined_v0.7.5.md`
- `docs\Release_Notes_0.7.4_to_0.7.5.md`
- `documents\CleanupSuite_Human_Friendly_User_Manual_v0.7.5.docx`
- `documents\CleanupSuite_Programmers_Guide_Adding_Tools_v0.7.5.docx`
- `AI_HANDOFF.md`

Summary of changes:
- Renamed the 0.7.4 documentation add-on files to 0.7.5 release-documentation files.
- Added `docs\Release_Notes_0.7.4_to_0.7.5.md`.
- Updated README and VERSIONING to describe `0.7.5` as the current documented release based on locked `0.7.4` behavior.
- Updated `SUITE_VERSION` from `0.7.4` to `0.7.5`.
- Rebuilt `VBA_Cleanup_tool.txt`, refreshed the practice bundle, and synced the updated version module into both `.docm` files.
- Updated the version assertion test to expect `0.7.5`.

Important observations:
- This is not a feature release. Tool behavior should remain the same as `0.7.4`.
- The `.docm` files changed because the embedded `modCleanupLauncher` version module was synced to `0.7.5`.
- An initial bulk text replacement caused encoding damage in a few existing files; those local edits were restored before final patching, and the final diffs for existing text/source files are targeted version-only changes.

Tests/checks run:
- Backup zip created and verified; it does not include `backups`.
- `python assemble.py` -> build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- `python -m unittest discover -s tests` -> 26 tests OK.
- `python vbaeval.py validate --file=VBA_Cleanup_tool.txt` -> `OK 24 builders | balance, if-traps, wiring, strings, round-trip all clean`.
- Source/generated/practice bundle check confirmed `0.7.5` version strings.
- Embedded `.docm` COM read-back confirmed both `.docm` files have `SUITE_VERSION = "0.7.5"` and not `0.7.4`.
- DOCX files were readable through bundled Python `python-docx` and their target-release text says `0.7.5 documentation-complete release`.
- DOCX render QA with bundled `render_docx.py` was attempted but failed because the renderer could not find LibreOffice/soffice (`FileNotFoundError [WinError 2]`).
- No Office `~$` lock files were found, and no `WINWORD` process remained after COM sync.

Known remaining issues:
- DOCX visual render QA is still blocked by the missing LibreOffice/soffice renderer dependency.
- Screenshots remain intentionally deferred.

Next recommended step:
- Run final `git diff --check`, status, and tests after this handoff entry, then commit with `Prepare 0.7.5 documentation release` and push `codex-0.7.0-start`.

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

## Session: 2026-06-12 14:32
Agent: Codex

Goal:
- Reinstall Cleanup Suite from the current local working tree.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_14-27-31.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: clean relative to `origin/main` except Chris had locally edited the Preview Actions action-bar size in `src\forms\frmPreviewActions.bas` and `src\installer\installer.bas` from `530/160` to `300/90`.
- Status after: dirty working tree with the `300/90` size changes, rebuilt generated bundles, and both `.docm` files reinstalled/updated.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root bundle, generated practice bundle, main `.docm`, and practice `.docm`; all verified to contain `FORM_W = 300` and `FORM_H = 90` for Preview Actions.

Files changed:
- `src\forms\frmPreviewActions.bas`
- `src\installer\installer.bas`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Preserved Chris's local action-bar size edit (`300/90`) rather than reverting it to the previously pushed `530/160`.
- Rebuilt `VBA_Cleanup_tool.txt` from current source and synced the practice generated bundle.
- Refreshed `modCleanupSuiteInstaller` inside both `.docm` files from the rebuilt generated bundle.
- Attempted `ReinstallCleanupSuite`; the deferred `OnTime` install did not complete and left forms removed.
- Corrected the half-reinstalled state by running `InstallCleanupSuite` directly in both `.docm` files.
- Saved both `.docm` files after successful direct install.

Important observations:
- After the first deferred reinstall pass, both `.docm` files had only installer/helper modules plus empty `UserForm*` shells; `frmPreviewActions` and other tool forms were missing.
- Running `InstallCleanupSuite` directly restored all expected Cleanup Suite forms and modules.
- Verification found 28 VBA components in each `.docm`, all expected suite components present, no numbered duplicate suite components, blank `frmPreviewActions` design caption, 6 Preview Actions controls, and saved state true.
- The current `300/90` action-bar size contradicts the pushed test expectations for `530/160`; this may clip visually under high DPI unless Chris intentionally wants to test that smaller size.
- Word was closed after automation. Installer report Notepad windows were opened; two closed cleanly, one packaged Windows Notepad process remained with an encoded CleanupSuite install-report session path despite `Stop-Process`.

Tests/checks run:
- Created backup zip before reinstall work.
- Ran bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Ran bundled Python full test discovery: `python -m unittest discover -s tests` -> failed 2 tests because they still expect `FORM_W = 530` and `FORM_H = 160` in `frmPreviewActions` and installer source.
- Verified both `.docm` files read-only after direct install: expected components present, `FORM_W = 300`, `FORM_H = 90`, installer also has `300/90`, `PreviewControls = 6`, `DesignerCaption=[]`, no numbered duplicate suite components, saved state true.
- Confirmed no `WINWORD` process remained.

Known remaining issues:
- Tests need either the action-bar expectation updated to `300/90` or the action bar reverted to the last verified `530/160` size.
- The `300/90` action bar has not been visually checked after reinstall and may be too small at 200% scaling.
- One packaged Windows Notepad process for an install report remained after attempted cleanup.

Next recommended step:
- Decide whether `300/90` is the intended new size. If yes, update the tests and visually verify the action bar in Word; if not, restore `530/160`, rebuild, reinstall, and rerun tests.

## Session: 2026-06-12 14:52
Agent: Codex

Goal:
- Keep the Preview Actions bar at Chris's smaller `300/90` shell size, but reduce font pressure and make the buttons slightly larger.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_14-45-38.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree from the `300/90` local size edit, rebuilt bundles, reinstalled `.docm` files, and the prior handoff entry.
- Status after: dirty working tree with tests/source/generated bundles/embedded `.docm` copies updated for the adjusted small-bar style.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root bundle, generated practice bundle, main `.docm`, and practice `.docm`.

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
- Preserved the smaller Preview Actions shell: `FORM_W = 300`, `FORM_H = 90`.
- Reduced visual text pressure: tool name font `8`, summary font `7.5`, button font `7.5`.
- Made buttons slightly larger without overflowing the small shell: `BTN_W = 84`, `BH = 18`, `CONTENT_W = 262`.
- Updated tests to encode the new intended small-bar geometry and typography.
- Rebuilt generated bundles and updated both embedded `.docm` copies with the new `frmPreviewActions` code/control positions and refreshed installer module text.

Important observations:
- Chris's screenshots showed the window shell size was not the issue; the labels were too large for the buttons.
- A trial with `BTN_W = 90` overflowed under the automated `PrintWindow` capture, so the final button width was reduced to `84`.
- Automated `PrintWindow` capture appears unreliable/cropped for the very small `300/90` form; verification relied on code/embedded-state checks plus Chris's real screenshots for the visual diagnosis.

Tests/checks run:
- Created backup zip before edits.
- TDD red check: focused preview test failed after updating expectations because production still had 80-wide buttons and larger fonts.
- Focused test after production changes: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Full test discovery: `python -m unittest discover -s tests` -> 14 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Verified both `.docm` files read-only: `FORM_W = 300`, `FORM_H = 90`, `BTN_W = 84`, `BH = 18`, smaller font calls present, installer matches, 6 controls, blank design caption, no temp audit modules, saved state true.
- Confirmed no `WINWORD` or `notepad` process remained.
- Removed temporary capture screenshots.

Known remaining issues:
- The adjusted small bar should still be checked by Chris in the live Word UI, because automated capture is unreliable for this very small form.
- The working tree is dirty and has not been committed/pushed after this adjustment.

Next recommended step:
- Have Chris visually confirm the new small-bar button/font balance in Word. If it looks right, commit and push this adjusted style.

## Session: 2026-06-12 15:04
Agent: Codex

Goal:
- Fix Preview Actions moving back to its startup position when the user toggles preview off after dragging the bar.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_15-00-45.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree from the small-bar style adjustment.
- Status after: dirty working tree with position-preservation fix added to source, tests, generated bundles, and both embedded `.docm` files.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `src\forms\frmPreviewActions.bas`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Root cause: `cmdPreview_Click` toggled preview off by calling `Me.Hide` and then `Me.Show vbModeless`; MSForms reapplied the form startup position during the second show.
- Added `ShowModelessAtCurrentPosition`, which stores `Me.Left` and `Me.Top`, hides/shows the form modelessly, then restores the saved coordinates.
- Updated the preview-off toggle path to call the helper instead of directly hiding/showing.
- Added regression assertions requiring the helper and coordinate restoration.
- Rebuilt generated bundles and updated both embedded `.docm` copies with the refreshed `frmPreviewActions` code and installer module text.

Important observations:
- This fix preserves the user's current dragged position only across the Preview ON -> OFF toggle in the same panel instance.
- It does not try to persist panel position across a full close/reopen or a new Preview Actions instance.

Tests/checks run:
- Created backup zip before edits.
- TDD red check: focused preview test failed because the position-preserving helper was not present.
- Focused test after code change: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Full test discovery: `python -m unittest discover -s tests` -> 14 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Verified both `.docm` files read-only: position helper present, toggle uses helper, no temp audit modules, saved state true.
- Confirmed no `WINWORD` or `notepad` process remained.

Known remaining issues:
- Chris should confirm in live Word that dragging the bar, then clicking `Preview is ON`, leaves the bar in place.
- The working tree is dirty and has not been committed/pushed after the small-bar style plus position fix.

Next recommended step:
- Live-test the drag-and-toggle behavior in Word. If it feels right, commit and push the accumulated small-bar style and position-preservation changes.

## Session: 2026-06-12 15:18
Agent: Codex

Goal:
- Fix Preview Actions moving back to its startup position when the user turns preview back on after turning it off.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_15-00-45.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree from the small-bar style adjustment and preview-off position fix.
- Status after: dirty working tree with preview-on position memory added to source, tests, generated bundles, and both embedded `.docm` files.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `src\modules\modCleanupHelpers.bas`
- `src\forms\frmPreviewActions.bas`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Root cause: the preview-on path unloads the current Preview Actions form and asks the source tool to run `PreviewFromPanel`, which creates a new Preview Actions instance through `ShowPreviewActions`; the new instance did not know where the user had dragged the old one.
- Added shared preview-panel position memory in `modCleanupHelpers`: `RememberPreviewActionPanelPosition` stores `Left/Top`, and `ApplyPreviewActionPanelPosition` sets `StartUpPosition = 0` plus the stored `Left/Top` before showing the newly-created panel.
- Updated `frmPreviewActions.cmdPreview_Click` so the preview-on branch remembers `Me.Left` and `Me.Top` before unloading and calling `PreviewFromPanel`.
- Kept the prior preview-off helper (`ShowModelessAtCurrentPosition`) for the same-instance hide/show path.
- Rebuilt generated bundles and updated both embedded `.docm` copies with refreshed `modCleanupHelpers`, `frmPreviewActions`, and installer module text.

Important observations:
- Position is now preserved across both preview toggles:
  - ON -> OFF: same panel hides/shows and restores its own position.
  - OFF -> ON: old panel stores position; the newly-created panel applies it before show.
- Position memory is reset by `ResetCleanupSuiteDefaults`.
- This is intentionally session-level behavior, not a persistent saved preference across future Word sessions.

Tests/checks run:
- TDD red check: focused preview test failed because shared position memory helpers were absent.
- Focused test after code changes: `python -m unittest tests.test_preview_action_panel` -> 10 tests OK.
- Full test discovery: `python -m unittest discover -s tests` -> 14 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Verified both `.docm` files read-only: helper stores position, helper applies before `gPreviewActionPanel.Show`, preview form remembers position before `PreviewFromPanel`, no temp audit modules, saved state true.
- Confirmed no `WINWORD` or `notepad` process remained.

Known remaining issues:
- Chris should live-test moving the bar, toggling preview off, then toggling preview back on to confirm the panel stays put through both transitions.
- The working tree is dirty and has not been committed/pushed after the small-bar style plus both position fixes.

Next recommended step:
- If Chris confirms the bar now stays where placed through both toggles, commit and push the accumulated changes.

## Session: 2026-06-12 15:52
Agent: Codex

Goal:
- Update the Main Menu wording/alignment, add a global "Return to main menu after closing individual tool menus" setting, and remove the launcher design-time caption repeat.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_15-46-16.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree from the compact Preview Actions bar and position-preservation changes.
- Status after: dirty working tree with the Main Menu changes added to source, tests, generated bundles, and both embedded `.docm` files.

Files changed:
- `src\forms\frmCleanupSuiteLauncher.bas`
- All individual tool form files under `src\forms\frm*.bas` except `frmCleanupSuiteLauncher.bas` and `frmPreviewActions.bas`
- `src\modules\modCleanupHelpers.bas`
- `src\installer\installer.bas`
- `tests\test_launcher_redesign.py`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `AI_HANDOFF.md`

Summary of changes:
- Shortened the Main Menu subtitle to: "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."
- Centered `lblLauncherTitle` and `lblLauncherSubtitle` in both runtime launcher code and installer-created designer styling.
- Added `chkReturnToMainAfterClose` between the existing "Return to main menu after apply" and "Auto-save before running each tool" global settings.
- Added persistent document-property helpers: `GetReturnToMainAfterCloseSetting`, `SetReturnToMainAfterCloseSetting`, and `ReturnToMainAfterToolClose`.
- Reset defaults now restores the close-return setting to `True`.
- Added `UserForm_QueryClose` to every individual tool menu so closing with the window close button reopens the launcher only when the new setting is enabled.
- Blank design-time caption is applied in `LayoutLauncherControls`; runtime `frmCleanupSuiteLauncher.UserForm_Initialize` sets `Me.Caption = "Cleanup Suite"` so the title bar remains useful without the in-body duplicate-caption artifact.

Tests/checks run:
- TDD red check: `python -m unittest tests.test_launcher_redesign` failed for the new missing launcher setting/close hooks before production code was changed.
- Focused tests after code changes: bundled Python `-m unittest tests.test_launcher_redesign tests.test_preview_action_panel` -> 16 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Updated both embedded `.docm` files with the assembled installer module, refreshed helper/form modules, and the new launcher checkbox/control layout.
- Embedded read-back verified both `.docm` files contain the generated installer builder, the new checkbox, centered title/subtitle, the shortened subtitle, launcher close-setting code, and individual tool close hooks.
- Final full test discovery after generated/docm/handoff updates: bundled Python `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- A live Word UI smoke test of closing a tool menu with the new setting on/off would be useful before the next commit.

## Session: 2026-06-12 19:55
Agent: Codex

Goal:
- Prepare `0.6.4` as the final programming-side release candidate / code-freeze target, then commit and push it for documentation work.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_19-47-28_v0.6.4-code-freeze.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: dirty working tree containing the compact Preview Actions bar, position-preservation fixes, Main Menu setting/layout work, generated bundles, embedded `.docm` updates, and tests.
- Status after: pending commit/push of the `0.6.4` code-freeze candidate.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `README.md`
- `VBA_Cleanup_tool.txt`
- `VERSIONING.md`
- all individual tool form source files under `src\forms\frm*.bas` except `frmPreviewActions.bas` changes are from the compact preview/action work and `frmCleanupSuiteLauncher.bas` changes are from the Main Menu setting/layout work
- `src\forms\frmPreviewActions.bas`
- `src\installer\installer.bas`
- `src\modules\modCleanupHelpers.bas`
- `src\modules\modCleanupLauncher.bas`
- `tests\test_launcher_redesign.py`
- `tests\test_preview_action_panel.py`

Summary of changes:
- Bumped `SUITE_VERSION` and project milestone metadata from `0.6.0` to `0.6.4`.
- Documented `0.6.4` as the code-freeze / final programming-side release candidate.
- Documented `0.6.5` as the documentation-complete release using the same codebase, except for documentation-blocking bug fixes if required.
- Preserved the already-completed programming changes in this local candidate: compact Preview Actions bar, preview-panel position preservation, Main Menu subtitle/centering, close-to-main-menu setting, and moved-up Main Menu global settings/reset controls.
- Rebuilt generated bundles and refreshed embedded `.docm` modules so generated/distributable and embedded states carry the same `0.6.4` version.

Important observations:
- After `0.6.4` is pushed, no feature/code changes should go into `0.6.5` unless they are required to correct documentation-blocking bugs.
- Do not start `0.7.0` work yet.

Tests/checks run:
- Created fresh milestone backup zip and verified it exists, excludes the `backups/` folder, and `backups/` is ignored by Git.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Embedded read-back verified both `.docm` files contain `SUITE_VERSION = "0.6.4"`, the assembled installer has `0.6.4`, installer builder functions are present, and both documents are saved.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- A live Word UI smoke test of the final `0.6.4` candidate remains useful, but this handoff marks the programming-side candidate ready for commit/push.

Next recommended step:
- Commit and push `0.6.4` to GitHub, then pause programming work while documentation is prepared from that pushed state.

## Session: 2026-06-12 20:18
Agent: Codex

Goal:
- Create the `0.6.5` documentation-complete release from the pushed `0.6.4` code-freeze state using ChatGPT's documentation package.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_20-09-48_v0.6.5-docs-release.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `main`
- Status before: clean at pushed `0.6.4` code-freeze commit `c787a70`.
- Status after: pending commit/push of the `0.6.5` documentation-complete release.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `README.md`
- `VBA_Cleanup_tool.txt`
- `VERSIONING.md`
- `src\modules\modCleanupLauncher.bas`
- `tests\test_launcher_redesign.py`
- `docs\User_Manual_v0.6.5.md`
- `docs\Quick_Start_v0.6.5.md`
- `docs\Troubleshooting_v0.6.5.md`
- `docs\Release_Notes_0.6.4_to_0.6.5.md`
- `docs\Documentation_Handoff_0.6.5.md`
- `docs\README_DOC_PACKAGE.md`
- `documents\CleanupSuite_User_Manual_v0.6.5.docx`

Summary of changes:
- Added the polished documentation package from `cleanup_suite_docs_0_6_5_from_0_6_4.zip`.
- Stored the package README as `docs\README_DOC_PACKAGE.md` to keep release-doc package notes with the other docs.
- Updated project version metadata from `0.6.4` to `0.6.5` in `README.md`, `VERSIONING.md`, `src\modules\modCleanupLauncher.bas`, and the matching version test.
- Rebuilt generated bundles so the distributable installer text carries `SUITE_VERSION = "0.6.5"`.
- Updated both embedded `.docm` files by replacing only `modCleanupLauncher` and the assembled installer module for the version bump.
- Confirmed no feature changes, UI changes, Preview Actions changes, launcher behavior changes, or cleanup-tool behavior changes were made beyond required version metadata and generated/docm updates caused by that metadata.

Important observations:
- `0.6.5` is documentation-complete relative to `0.6.4`; do not start `0.7.0` work until Chris confirms this release looks good.
- The documentation package states it was written from pushed `0.6.4` commit `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`.

Tests/checks run:
- Created fresh backup zip and verified it exists and excludes the `backups/` folder.
- Inspected the ChatGPT documentation package contents before copying files.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Embedded read-back verified both `.docm` files contain `SUITE_VERSION = "0.6.5"`, the assembled installer has `0.6.5`, installer builder functions are present, and both documents are saved.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- None for the documentation-complete release commit. Live review of the new docs by Chris is still the safest next human check.

Next recommended step:
- Commit and push `0.6.5`, then wait for Chris confirmation before any `0.7.0` work.

## Session: 2026-06-12 21:30
Agent: Codex

Goal:
- Start the `0.7.0` development line cleanly after Chris confirmed the `0.6.5` documentation-complete release.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_21-18-46_v0.7.0-start.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Starting branch/status: clean `main` at pushed `0.6.5` commit `d8ea6a1`.
- Working branch: `codex-0.7.0-start` (slash-style `codex/0.7.0-start` branch creation hit a local Git ref-lock/permission issue, so this safe equivalent branch name was used).
- Status after: pending commit of the `0.7.0` starting-point metadata and generated/docm sync.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `README.md`
- `VBA_Cleanup_tool.txt`
- `VERSIONING.md`
- `src\modules\modCleanupLauncher.bas`
- `tests\test_launcher_redesign.py`

Summary of changes:
- Bumped project version metadata from `0.6.5` to `0.7.0`.
- Updated `VERSIONING.md` to define the next release rhythm:
  - `0.7.0` = next programming milestone begins.
  - `0.7.4` = code freeze / final programming-side release candidate for `0.7.x`.
  - `0.7.5` = documentation-complete release using the same codebase as `0.7.4`, except for documentation-blocking bug fixes if required.
- Rebuilt generated bundles and refreshed embedded `.docm` modules so generated/distributable and embedded states carry `SUITE_VERSION = "0.7.0"`.
- No feature work, UI work, Preview Actions changes, launcher behavior changes, or cleanup-tool behavior changes were made.

Important observations:
- This commit should be treated as the clean start marker for `0.7.0`; actual feature planning/implementation should happen after this marker.

Tests/checks run:
- Created fresh backup zip and verified it exists and excludes the `backups/` folder.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Embedded read-back verified both `.docm` files contain `SUITE_VERSION = "0.7.0"`, the assembled installer has `0.7.0`, installer builder functions are present, and both documents are saved.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- None for the `0.7.0` start marker.

Next recommended step:
- Commit this starting point, then decide the first scoped `0.7.0` feature task before making behavior changes.

## Session: 2026-06-12 21:45
Agent: Codex

Goal:
- Fix the Main Menu overlap where the moved-up global settings painted over the right-column Review section.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_21-33-48_v0.7.0-launcher-overlap-fix.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: clean at local `0.7.0` start commit `b2651c4`.
- Status after: pending commit of the launcher overlap fix.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `VBA_Cleanup_tool.txt`
- `src\installer\installer.bas`
- `tests\test_launcher_redesign.py`

Summary of changes:
- Root cause: the global setting checkboxes were moved upward under the left column, but their generated widths still spanned the full form (`FORM_W - 2 * MARGIN`). MSForms paints the checkbox background across that full rectangle, covering controls in the right column.
- Changed generated launcher layout so `chkReturnToMainAfterApply`, `chkReturnToMainAfterClose`, and `chkAutoSave` use `COL_W` width, confining their painted area to the left column.
- Added a regression test asserting the setting checkboxes use `COL_W` and do not use the full-form width.
- Rebuilt generated bundles and updated both embedded `.docm` copies.

Important observations:
- Window size was not increased.
- The right-column Review section should no longer be hidden by the global settings controls.

Tests/checks run:
- TDD red check: focused launcher test failed while the checkboxes still used full-form width.
- Focused launcher test passed after the fix.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Embedded read-back verified both `.docm` files have installer layout using `COL_W`, and the three setting checkbox widths match the left-column width (`358`).
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- Chris should visually re-open the Main Menu and confirm the Review section is no longer covered.

Next recommended step:
- Commit this overlap fix on `codex-0.7.0-start`; then continue with a scoped `0.7.0` feature plan.

## Session: 2026-06-12 22:02
Agent: Codex

Goal:
- Change the Preview Actions OFF-state message so it invites editing instead of repeating the preview findings summary.

Backup created:
- `backups\CleanupSuite_backup_2026-06-12_21-52-53_preview-off-message.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: clean at local overlap-fix commit `c8e040d`; an Office temp lock file from a visible/hidden Word session briefly appeared and was not treated as project content.
- Status after: pending commit of the Preview Actions OFF-message change.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `VBA_Cleanup_tool.txt`
- `src\forms\frmPreviewActions.bas`
- `tests\test_preview_action_panel.py`

Summary of changes:
- In `frmPreviewActions.cmdPreview_Click`, when preview is turned OFF the summary label now says: "You may now edit your document if you wish."
- The ON-state still shows the original tool findings summary through `Configure`.
- Added a regression test that verifies the OFF caption is set after `cmdPreview.Caption = "Preview is OFF"`.
- Rebuilt generated bundles and updated both embedded `.docm` copies with the refreshed Preview Actions form and assembled installer.

Important observations:
- No launcher changes, cleanup-tool behavior changes, or layout size changes were made for this request.

Tests/checks run:
- Created backup zip after resolving the file-lock problem with shared-read backup handling; skipped only Office `~$` temp files.
- TDD red check: focused preview-panel test failed before the source change.
- Focused preview-panel test passed after the source change.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Embedded read-back verified both `.docm` files contain the new OFF message in `frmPreviewActions` and in the assembled installer module; both documents saved.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 16 tests OK.

Known remaining issues:
- Chris should live-test toggling Preview OFF to confirm the panel displays the edit message.

Next recommended step:
- Commit the Preview OFF message fix on `codex-0.7.0-start`.

## Session: 2026-06-13 01:26
Agent: Codex

Goal:
- Begin the 0.7.0 sub-tool menu redesign with Capitalization Fixer as the first guided, more modern form.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_00-58-49_capitalization-guided-form.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: clean after `4b1ac15 Update preview off message`.
- Status after: pending source/test/generated changes; embedded `.docm` files not updated.
- GitHub/local/generated/.docm sync checked: local source and generated bundles synced; embedded `.docm` sync attempted but blocked by Word COM timeouts.

Files changed:
- `AI_HANDOFF.md`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `src\forms\frmCapitalizationCleanup.bas`
- `src\installer\installer.bas`
- `src\modules\modCleanupHelpers.bas`
- `tests\test_capitalization_guided_form.py`
- `tests\test_preview_action_panel.py`

Summary of changes:
- Reworked Capitalization Fixer into a guided mode-selection flow while preserving the existing transformation logic.
- Added `lblIntro` and `lblModeSummary` so the form explains the selected action instead of presenting only mechanical controls.
- Renamed mode captions to user-facing actions: `Fix sentence starts`, `Sentence case`, `Title case`, `UPPERCASE`, `lowercase`, and `Custom`.
- Added dynamic summary text and click handlers so the guidance updates when a mode changes.
- Added `LayoutCapitalizationCleanupForm` for a compact two-column guided layout, with Custom controls hidden until Custom is selected.
- Updated the installer control list, captions, style handling, and generated design-time layout for the new Capitalization controls.
- Updated tests so Capitalization may use its specialized layout while other preview-capable forms continue using the shared layout helper.

Important observations:
- No capitalization cleanup algorithm changes were made. The existing `optAll`, `optSentence`, `optTitle`, `optUpper`, `optLower`, and `optCustom` behavior remains wired the same way.
- Generated root and practice `VBA_Cleanup_tool.txt` files compare equal.
- Multiple Word COM attempts to update/read the `.docm` files timed out or produced a Word installer/modal state. Stale Word processes and `~$` lock files were cleared afterward.
- Git does not show either `.docm` as modified, so Chris should treat the source/generated bundle as updated but the embedded documents as not yet synced for this change.

Tests/checks run:
- Focused tests: bundled Python `-m unittest tests.test_capitalization_guided_form tests.test_preview_action_panel` -> 12 tests OK.
- Full test discovery: bundled Python `-m unittest discover -s tests` -> 18 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.

Known remaining issues:
- Embedded `.docm` files still need a successful Word/VBA sync pass before Chris can test the new Capitalization form from the workspace documents.

Next recommended step:
- Resolve the Word COM `.docm` sync blocker, then update both `.docm` files and visually inspect Capitalization Fixer in Word.

## Session: 2026-06-13 08:30
Agent: Codex

Goal:
- Resolve the `.docm` sync blocker from the Capitalization guided-form pass and verify the embedded documents.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_08-12-42_before-docm-sync-retry.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: pending Capitalization guided-form source/test/generated changes; embedded `.docm` files not yet synced.
- Status after: pending Capitalization guided-form changes with both embedded `.docm` files updated and verified.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`

Summary of changes:
- Diagnosed the Word COM blocker: `Documents.Open` hung when using the long project path with spaces, but opened normally when using a temporary copy or the DOS short path.
- Updated both embedded `.docm` files by copying each document to `%TEMP%`, opening the temp copy, syncing changed VBA/form components, saving it, then replacing the original document with the synced copy.
- Synced embedded `modCleanupHelpers`, `frmCapitalizationCleanup`, and `modCleanupSuiteInstaller`.
- Added the new embedded `frmCapitalizationCleanup` label controls `lblIntro` and `lblModeSummary` to both documents.

Important observations:
- The actual `.docm` files are now modified in Git and contain the new Capitalization guided-form code and controls.
- Future Word COM automation against this project may need DOS short paths or temp-copy editing for reliable `.docm` opens.
- No Word processes or Office `~$` lock files remained after the sync.

Tests/checks run:
- Read-back verification using DOS short paths confirmed both `.docm` files contain:
  - `LayoutCapitalizationCleanupForm` in `modCleanupHelpers`
  - `lblIntro.Caption` and the custom summary text in `frmCapitalizationCleanup`
  - `LayoutCapitalizationControls` in `modCleanupSuiteInstaller`
  - embedded controls `lblIntro` and `lblModeSummary`
  - saved state true
- Runtime audit on a temporary synced `.docm` copy loaded `frmCapitalizationCleanup` successfully and recorded:
  - intro: `Choose how you want capitalization cleaned up.`
  - default summary: `Recommended: capitalizes likely sentence starts after . ? ! while leaving existing casing alone where possible.`
  - default mode: `Fix sentence starts: True`
  - custom visible: `False`
  - buttons: `Preview` and `Apply`
  - form size: `428 x 264`
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 18 tests OK.

Known remaining issues:
- Chris should visually open Capitalization Fixer in Word and decide whether this is the right first-pass feel before applying the same pattern to other sub-tool menus.

Next recommended step:
- Visually review Capitalization Fixer in the workspace `.docm`; then either refine this first guided-form pattern or choose the next tool menu to redesign.

## Session: 2026-06-13 09:48
Agent: Codex

Goal:
- Polish the first Capitalization guided-form pass based on Chris's screenshots and feedback.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_09-40-15_capitalization-polish.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: pending Capitalization guided-form source/test/generated/docm changes.
- Status after: pending Capitalization guided-form polish changes with generated bundles and both embedded `.docm` files updated.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root/practice bundles, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `VBA_Cleanup_tool.txt`
- `src\forms\frmCapitalizationCleanup.bas`
- `src\installer\installer.bas`
- `src\modules\modCleanupHelpers.bas`
- `tests\test_capitalization_guided_form.py`

Summary of changes:
- Removed the visible duplicate in-form title by adding runtime hiding for legacy controls whose caption is `Capitalization Fixer`.
- Added bottom room to the Capitalization guided layout so the bottom action buttons are not clipped.
- Fixed Custom mode visibility by sending the decorative `fraCustom` frame behind the actual custom checkbox controls and bringing choices/buttons forward.
- Cleaned and realigned the Help message, removing the awkward split sentence and switching to clearer mode-by-mode text.
- Updated installer-generated layout height/Z-order so fresh installs match the runtime polish.
- Added/updated regression coverage for the duplicate-title guard, bottom height, Custom Z-order, and Help text cleanup.

Important observations:
- No cleanup algorithm changes were made.
- The previous Word COM workaround was used successfully: update temp `.docm` copies, then replace the originals.
- No Word processes remained after verification.

Tests/checks run:
- Focused Capitalization guided-form tests: bundled Python `-m unittest tests.test_capitalization_guided_form` -> 2 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Embedded read-back using short paths confirmed both `.docm` files contain the duplicate-title hide helper, height padding, cleaned Help title, and installer height update; saved state true.
- Runtime audit on a temporary synced `.docm` copy loaded `frmCapitalizationCleanup` and recorded:
  - default height: `298`
  - default custom visible: `False`
  - default bottom button extent: `236`
  - Custom height: `456`
  - Custom frame visible: `True`
  - all five Custom choice checkboxes visible: `True`
  - Custom bottom button extent: `394`
  - summary: `Custom runs selected choices in order; later choices can override earlier ones.`
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 18 tests OK.

Known remaining issues:
- Chris should visually confirm the polished Capitalization form in Word. The automated runtime audit confirms the controls are visible and not clipped, but this final judgment is visual/design feel.

Next recommended step:
- Open Capitalization Fixer in Word, check normal and Custom modes, then decide whether to make one more polish pass or use this pattern for the next sub-tool menu.

## Session: 2026-06-13 10:24
Agent: Codex

Goal:
- Finish the remaining Capitalization guided-form issues from Chris's follow-up screenshots: Custom options visible and duplicate title removed.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_09-56-08_capitalization-custom-title-fix.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.7.0-start`
- Status before: pending Capitalization guided-form polish changes.
- Status after: pending Capitalization guided-form polish changes with source, generated bundles, tests, and both embedded `.docm` files updated.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root/practice bundles, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `VBA_Cleanup_tool.txt`
- `src\forms\frmCapitalizationCleanup.bas`
- `src\installer\installer.bas`
- `src\modules\modCleanupHelpers.bas`
- `tests\test_capitalization_guided_form.py`

Summary of changes:
- Changed Custom mode to show the five Custom checkboxes directly on the form instead of inside a visible `fraCustom` frame; the frame is now hidden at runtime and in generated layout.
- Confirmed the duplicate visible title was not a normal visible control in runtime audit; it was tied to the MSForms/UserForm caption behavior.
- Set `Me.Caption = ""` in `frmCapitalizationCleanup.UserForm_Initialize` and blanked the generated design-time caption for this one form to prevent the double-title display.
- Kept the cleaned Help dialog title/text intact.

Important observations:
- Help does not show the same double title because it is a native `MsgBox`, not the custom MSForms `UserForm` surface.
- The Capitalization form now relies on the centered instruction text rather than a native/userform caption.
- No cleanup algorithm changes were made.

Tests/checks run:
- Focused Capitalization guided-form tests: bundled Python `-m unittest tests.test_capitalization_guided_form` -> 2 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Updated both embedded `.docm` files using the temp-copy Word COM method.
- Runtime audit on a temporary synced `.docm` copy confirmed:
  - `formCaption=[]`
  - intro caption present
  - default height `298`
  - Custom frame hidden
  - Custom choices visible (`chkSentence=True`, `chkSmartSentences=True`)
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 18 tests OK.
- No Word processes or Office `~$` lock files remained after verification.

Known remaining issues:
- Chris should visually confirm the blank-caption approach removes the persistent duplicate title in the live Word form.

Next recommended step:
- Visually re-open Capitalization Fixer in Word in normal and Custom modes. If this is approved, commit the Capitalization guided-form pass or apply the same design pattern to the next sub-tool menu.

## Session: 2026-06-13 12:36
Agent: Codex

Goal:
- Apply the successful MSForms double-title handling as a shared project pattern, record both title-handling methods, and document project workarounds/coding mechanisms.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_12-18-12_projectwide-msform-caption-docs.zip`

Git/repo state:
- Branch: `codex-0.7.0-start`
- Status before: pending 0.7.0 Capitalization guided-form UI/source/generated/docm changes.
- Status after: pending shared MSForm title-strategy changes, generated bundle sync, embedded `.docm` sync, tests, and new workaround documentation.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed/added in this pass:
- `docs\Cool_work_arounds_and_coding_mechanisms_for_VBA_programmed_through_Python.md`
- `src\modules\modCleanupHelpers.bas`
- `src\forms\frmCleanupSuiteLauncher.bas`
- `src\forms\frmPreviewActions.bas`
- `src\forms\frmCapitalizationCleanup.bas`
- `src\installer\installer.bas`
- `tests\test_capitalization_guided_form.py`
- `tests\test_launcher_redesign.py`
- `tests\test_preview_action_panel.py`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`

Summary of changes:
- Added shared `ApplyMSFormTitleStrategy` helper with two recorded methods:
  - Method 1 blanks the native UserForm caption when the body owns the title.
  - Method 2 keeps the native caption and hides legacy in-body `lblTitle` controls.
- Applied Method 1 to the Main Menu, Preview Actions action bar, and Capitalization guided form.
- Applied Method 2 through `LayoutCleanupToolForm` for the older generic tool forms.
- Added installer-side `FormBodyOwnsTitle` so generated installs blank design-time captions for the body-owned-title forms.
- Added the requested workaround/coding-mechanism document under `docs\`.
- Rebuilt generated bundles and synced both embedded `.docm` files.

Important observations:
- Help dialogs do not show the same double-title issue because they are native `MsgBox` dialogs, not custom MSForms UserForms.
- The failed long-path read-back attempt reproduced the known Word COM limitation; verification was redone successfully from short temp copies.
- One leftover invisible Word automation process from the timed-out long-path read-back was closed/cleared; no Word processes remained after final checks.
- No cleanup algorithm changes were made as part of this title/documentation pass.

Tests/checks run:
- Focused tests after source updates: bundled Python `-m unittest tests.test_capitalization_guided_form tests.test_launcher_redesign tests.test_preview_action_panel` -> 18 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Updated both embedded `.docm` files using the temp-copy Word COM method.
- Embedded read-back from short temp copies confirmed both `.docm` files contain `ApplyMSFormTitleStrategy`, `HideLegacyTitleControls`, `FormBodyOwnsTitle`, and `ApplyMSFormTitleStrategy Me, True` in Launcher, Preview Actions, and Capitalization.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 18 tests OK.

Known remaining issues:
- Chris should visually confirm the Main Menu, Preview Actions bar, Capitalization form, and one generic older sub-tool form in Word to make sure the two title methods feel right in the actual UI.

Next recommended step:
- If the shared title behavior looks good visually, continue the 0.7.0 guided-form modernization with the next sub-tool menu.

## Session: 2026-06-13 13:16
Agent: Codex

Goal:
- Fix runtime error 400 when reopening the same tool from the Main Menu after returning to it.
- Add a global completion-review setting and clarify how it interacts with "Return to main menu after apply."

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_13-08-47_return-review-reopen-fix.zip`

Git/repo state:
- Branch: `codex-0.7.0-start`
- Status before: pending 0.7.0 guided-form/title-strategy work plus generated/docm sync.
- Status after: pending reopen fix, completion-review setting, generated bundle sync, embedded `.docm` sync, tests, and this handoff entry.
- GitHub/local/generated/.docm sync checked: yes for local source, root generated bundle, practice generated bundle, main `.docm`, and practice `.docm`.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `VBA_Cleanup_tool.txt`
- all individual tool form files under `src\forms` for the close callback update
- `src\forms\frmCleanupSuiteLauncher.bas`
- `src\installer\installer.bas`
- `src\modules\modCleanupHelpers.bas`
- `tests\test_launcher_redesign.py`
- `tests\test_preview_action_panel.py`

Summary of changes:
- Root cause of error 400: the launcher was trying to show a default UserForm instance that could still be considered displayed after returning to the Main Menu from that same tool.
- Added `OpenCleanupTool(toolForm As Object)` in the launcher. It hides the launcher, hides the target tool instance if it is already displayed, then shows it. Tool buttons now call this helper instead of direct `frmX.Show`.
- Updated individual tool `UserForm_QueryClose` handlers to call `ReturnToMainAfterToolClose Me`, allowing the helper to hide the closing form before showing the Main Menu.
- Updated `ReturnToMainAfterToolClose` to accept the optional closing form and hide it before showing the launcher.
- Added global setting `Show completion review after apply`, stored as `CleanupSuiteShowCompletionReviewAfterApply`.
- Default for completion review is ON.
- `ShowCleanupReport` now exits silently when completion review is disabled.
- Main Menu wording now changes dynamically:
  - review ON: `Return to main menu after completion review`
  - review OFF: `Return to main menu after apply`
- Reset All restores completion review to ON.

Important observations:
- The new completion-review setting controls whether post-apply result dialogs appear.
- The return-to-main setting controls where the user lands after apply/review flow finishes.
- No cleanup algorithms were changed.
- Word/VBE was not running during `.docm` sync; no Word processes remained after verification.

Tests/checks run:
- Focused tests initially failed for the expected missing behavior, then passed after implementation.
- Focused tests: bundled Python `-m unittest tests.test_launcher_redesign tests.test_preview_action_panel` -> 17 tests OK.
- Full tests before generated sync: bundled Python `-m unittest discover -s tests` -> 19 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice generated bundles compare equal.
- Updated both embedded `.docm` files using the temp-copy Word COM method.
- Embedded read-back from short temp copies confirmed both `.docm` files contain:
  - `OpenCleanupTool`
  - `toolForm.Hide` before `toolForm.Show`
  - Capitalization launcher click uses `OpenCleanupTool frmCapitalizationCleanup`
  - Capitalization close callback passes `Me`
  - completion-review helper functions
  - `ShowCleanupReport` skip guard
  - installer-generated completion-review control
- Final full test discovery: bundled Python `-m unittest discover -s tests` -> 19 tests OK.

Known remaining issues:
- Chris should visually confirm the exact sequence that failed before: open a tool, return to Main Menu, reopen the same tool. The source/generated/embedded checks specifically cover the fix, but the final confidence check is in live Word.

Next recommended step:
- Test the Main Menu settings in Word:
  - with completion review ON, confirm the return label says `Return to main menu after completion review`;
  - with completion review OFF, confirm it says `Return to main menu after apply`;
  - confirm reopening the same tool no longer raises runtime error 400.

## Session: 2026-06-13 13:28
Agent: Codex

Goal:
- Fix compile error when launching CleanupSuite after adding the `Show completion review after apply` setting.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_13-22-17_launcher-missing-review-control-fix.zip`

Git/repo state:
- Branch: `codex-0.7.0-start`
- Status before: pending reopen/completion-review source/generated/docm changes.
- Status after: same pending source/generated/docm changes, with both embedded `.docm` launcher designers patched to include the missing checkbox control.
- GitHub/local/generated/.docm sync checked: yes for both embedded `.docm` files.

Files changed:
- `AI_HANDOFF.md`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`

Summary of changes:
- Root cause: source code and generated installer knew about `chkShowCompletionReviewAfterApply`, but the existing embedded `.docm` form designer did not have that new checkbox control. VBA therefore raised `Compile error: Variable not defined` in `frmCleanupSuiteLauncher.UserForm_Initialize`.
- Patched both embedded `frmCleanupSuiteLauncher` designers to add `chkShowCompletionReviewAfterApply`.
- Positioned the new checkbox between `chkReturnToMainAfterApply` and `chkReturnToMainAfterClose`, then moved `chkReturnToMainAfterClose`, `chkAutoSave`, and `cmdResetAll` down to match the installer layout.

Important observations:
- This was not a source-code bug in the installer path; it was an embedded `.docm` designer-sync bug.
- When adding a new MSForms control, syncing only code modules is not enough. The existing `.docm` designer must also receive the new control, or the form code will not compile.
- First COM designer patch attempt failed due to a stale/null object reference from `Controls.Add`; retrying by adding the control and then fetching it by name was reliable.
- One invisible Word automation instance was left after the failed COM attempt and had to be force-stopped because it rejected cleanup calls.

Tests/checks run:
- Patched both `.docm` files using Word COM and short temp copies.
- Embedded read-back verified both `.docm` launchers contain `chkShowCompletionReviewAfterApply` with caption `Show completion review after apply`.
- Smoke macro on short temp copies created and unloaded `frmCleanupSuiteLauncher` successfully in both `.docm` files, proving the compile error path is fixed.
- Full test discovery with bundled Python: `-m unittest discover -s tests` -> 19 tests OK.
- No Word processes remained after verification.

Known remaining issues:
- Chris should retry launching from the Ribbon button and Alt-F8 `ShowCleanupSuiteLauncher` in the live document to confirm the compile dialog is gone.

Next recommended step:
- After visual/runtime confirmation, consider improving the `.docm` sync workflow so new controls are automatically created or the installer is rerun whenever `ControlsForForm` changes.

## Session: 2026-06-13 13:58
Agent: Codex

Goal:
- Fix the runtime error sequence where reopening a tool from the Main Menu, using Help/Apply, and closing nested forms could leave stale form state and raise `Run-time error '13': Type mismatch`.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_13-42-26_byval-preview-flag-fix.zip`

Git/repo state:
- Branch: `codex-0.7.0-start`
- Status before: pending 0.7.0 form-redesign/global-setting changes.
- Status after: same pending 0.7.0 changes, plus the launcher/tool reopen fix and preview-flag reset.
- Generated bundle and both `.docm` files synced after source changes.

Files changed in this fix:
- `src\forms\frmCleanupSuiteLauncher.bas`
- `src\modules\modCleanupHelpers.bas`
- all preview-capable `src\forms\frm*.bas` tool forms
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`
- `tests\test_launcher_redesign.py`
- `tests\test_preview_action_panel.py`
- `tests\test_capitalization_guided_form.py`
- `AI_HANDOFF.md`

Summary of changes:
- Root cause 1: `OpenCleanupTool(toolForm As Object)` passed UserForm default instances by reference. In this reopen path, VBA could throw `Run-time error '13': Type mismatch`. The launcher helper now uses `Private Sub OpenCleanupTool(ByVal toolForm As Object)`.
- Root cause 2: `PreviewFromPanel` set `chkPreviewOnly.Value = True` before calling `cmdRun_Click`, but did not clear it afterward. Returning to the sub-tool and pressing Apply could therefore launch Preview Actions again instead of the apply/completion flow. Each preview-capable form now resets `chkPreviewOnly.Value = False` after `cmdRun_Click`.
- Related object-helper parameters in `modCleanupHelpers.bas` were made `ByVal` where they receive form/control objects.
- Removed a stray BOM marker from `frmCapitalizationCleanup.bas` before final rebuild/sync.

Tests/checks run:
- Focused tests after test updates: bundled Python `-m unittest tests.test_launcher_redesign tests.test_preview_action_panel tests.test_capitalization_guided_form` -> 19 tests OK.
- Full tests before sync: bundled Python `-m unittest discover -s tests` -> 19 tests OK.
- Rebuilt generated bundle with `python assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice SHA-256 hashes match.
- Synced both embedded `.docm` files using short temp copies and Word COM.
- Embedded read-back/smoke on both `.docm` files confirmed:
  - `OpenCleanupTool(ByVal toolForm As Object)`
  - `ReturnToMainAfterToolClose(Optional ByVal closingForm As Object = Nothing)`
  - capitalization `PreviewFromPanel` resets `chkPreviewOnly.Value = False`
  - `chkShowCompletionReviewAfterApply` exists in the launcher designer
  - temporary smoke macro can create and unload `frmCleanupSuiteLauncher`
- Final full tests after sync: bundled Python `-m unittest discover -s tests` -> 19 tests OK.
- No Word processes remained after verification.

Known remaining issues:
- Chris should retry the exact live Word path that failed: open a tool, use Help, close Help, Apply, close/reopen through Main Menu, and close the sub-tool. Source/generated/embedded checks cover the identified root causes, but the final confidence check is the live UI sequence.

## Session: 2026-06-13 14:22
Agent: Codex

Goal:
- Fix the no-visible-error runtime flaw where a sub-tool could be opened, closed, reopened, and then refuse to close cleanly, leaving a hidden/background submenu and confused Preview Actions/apply flow.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_14-15-34_tool-session-lifecycle-fix.zip`

Root cause:
- The previous close flow showed the Main Menu from inside each sub-tool's `UserForm_QueryClose` event. That meant a new modal form could be opened before the closing sub-tool had actually unwound, leaving hidden/default UserForm instances in an unstable state.
- The launcher also reused default UserForm instances (`frmCapitalizationCleanup`, etc.). That made stale hidden forms more likely after close/reopen cycles.

Summary of changes:
- The launcher now owns the whole tool session:
  - hides itself,
  - creates a fresh tool instance with `VBA.UserForms.Add(formName)`,
  - waits for `toolForm.Show` to return,
  - then decides whether to show the Main Menu again.
- Sub-tool `QueryClose` handlers no longer show the Main Menu directly. They call `ReturnToMainAfterToolClose Me, Cancel, CloseMode`, which only marks that the user closed the tool.
- Successful apply paths now call `MarkCleanupToolApplied` before unloading.
- Preview Actions no longer shows the Main Menu directly after Apply; it lets the launcher finish the session and make the return decision.
- Generated installer bundle, practice bundle, and both `.docm` files were synced.

Tests/checks run:
- RED tests first failed against the old lifecycle as expected.
- Focused lifecycle tests after implementation: bundled Python `-m unittest tests.test_launcher_redesign tests.test_preview_action_panel` -> 17 tests OK.
- Full tests before sync: bundled Python `-m unittest discover -s tests` -> 19 tests OK.
- Rebuilt generated bundle with `python assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`; root/practice SHA-256 hashes match.
- Synced both embedded `.docm` files using short temp copies and Word COM.
- Embedded Word COM smoke on both `.docm` temp copies verified:
  - `VBA.UserForms.Add("frmCapitalizationCleanup")` compiles and runs,
  - close-session marker returns to launcher when close setting is ON,
  - apply-session marker returns to launcher when apply setting is ON,
  - `frmCleanupSuiteLauncher` can instantiate/unload cleanly,
  - embedded code contains the new launcher and close-hook patterns.
- Final full test discovery: bundled Python `-m unittest discover -s tests` -> 19 tests OK.
- No Word processes remained after verification.

Known remaining issues:
- The exact UI sequence should still be visually retried in Word: open a tool, close it, open it again or open another tool then the first, close it, apply, preview/apply, and confirm no submenu remains behind the current workflow.

## Session: 2026-06-13 15:19
Agent: Codex

Goal:
- Apply the new 0.7.0 guided submenu direction across the rest of the cleanup-tool submenus while preserving the stable launcher/tool lifecycle fix.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_15-05-57_projectwide-guided-submenus.zip`

Summary of changes:
- Extended the shared generic submenu helper `LayoutCleanupToolForm` so all non-Capitalization tool forms inherit a cleaner guided layout:
  - body-owned runtime title and intro labels,
  - native MSForm caption blanking through `ApplyMSFormTitleStrategy toolForm, True`,
  - hidden frame placeholders,
  - consistent option/check/label styling,
  - compact scope row,
  - full-width Preview button,
  - Apply / Reset / Help bottom row.
- Added the matching generated-installer layout helper `LayoutGenericToolControls` so regenerated installs reserve the same guided surface and action layout.
- Expanded `FormBodyOwnsTitle` so all Cleanup Suite UserForms can avoid the duplicate-title problem through the body-title method.
- Preserved the older title-removal mechanism (`HideLegacyTitleControls` / `IsLegacyTitleControlName`) as a fallback and documented both approaches.
- Updated `docs\Cool_work_arounds_and_coding_mechanisms_for_VBA_programmed_through_Python.md` with:
  - the project-wide guided submenu mechanism,
  - the updated title strategy,
  - the completion-review global setting.
- Rebuilt `VBA_Cleanup_tool.txt` and synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.

Tests/checks run:
- Focused RED test before implementation failed as expected for missing project-wide guided layout helpers.
- Focused test after implementation: bundled Python `-m unittest tests.test_capitalization_guided_form` -> 3 tests OK.
- Full tests before assemble: bundled Python `-m unittest discover -s tests` -> 20 tests OK.
- Rebuilt generated bundle with `python assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Full tests after assemble and practice bundle sync: bundled Python `-m unittest discover -s tests` -> 20 tests OK.
- Final full test run after `.docm` sync/smoke: bundled Python `-m unittest discover -s tests` -> 20 tests OK.

Sync status:
- Source files and generated text bundles are synced.
- Embedded `.docm` sync was retried successfully using the normal short-temp-copy Word COM plan.
- Both `CleanupSuite_Workspace.docm` and `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm` were synced from repo source and verified by embedded read-back.
- Embedded read-back confirmed:
  - `EnsureGuidedLabel`,
  - `LayoutGuidedActionButtons toolForm, M, y, contentW`,
  - launcher `VBA.UserForms.Add(formName)`.
- Runtime smoke on a temporary synced `.docm` copy instantiated:
  - `frmCleanupSuiteLauncher` -> caption blank, size `767.9 x 553`,
  - `frmCapitalizationCleanup` -> caption blank, size `428 x 298`,
  - `frmUnicodeCleanup` -> guided title `Invisible Unicode Cleaner`,
  - `frmDuplicateDetector` -> guided title `Duplicate Paragraph Detector`,
  - `frmHeaderFooterStandardizer` -> guided title `Header / Footer Standardizer`,
  - `frmObjectRemover` -> guided title `Object Remover`,
  - `frmPreviewActions` -> caption blank, size `308 x 89.95`.
- No Word processes or Office `~$` lock files remained after verification.

Known remaining issues:
- Live visual QA is still useful, especially for high-control-count forms such as Header / Footer Standardizer and Object Remover. Runtime smoke confirms they instantiate with the guided title/introduction and expected compact widths.

## Session: 2026-06-13 22:50
Agent: Codex

Goal:
- Finish the two-column guided layout follow-up: generic tool choices and any Custom-revealed choices should follow the Capitalization tool's compact two-column pattern.
- Record that rule in the new-tool/building guidance so future tools do not drift back to single-column menus.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_22-34-31_before-two-column-guided-tools.zip`

Summary of changes:
- Updated `LayoutCleanupToolForm` in `src\modules\modCleanupHelpers.bas` so visible `opt`/`chk` choices are paired left/right into two-column rows.
- Applied the same two-column pairing rule to Custom-revealed choices, break conversion choices, duplicate fuzzy choices, and other conditional choice groups because they flow through the same shared layout helper.
- Updated `LayoutGenericToolControls` in `src\installer\installer.bas` so newly installed/rebuilt generic tools use the same two-column choice layout from the start.
- Kept command buttons, description labels, frames, hidden controls, and scoped controls in the existing guided flow; only the choice rows changed.
- Documented the two-column rule in:
  - `CONTRIBUTING.md`,
  - `README.md`,
  - `docs\Cool_work_arounds_and_coding_mechanisms_for_VBA_programmed_through_Python.md`.
- Rebuilt `VBA_Cleanup_tool.txt` and synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.

Tests/checks run:
- Focused RED tests first failed against the old single-column expectation.
- Focused test after implementation: bundled Python `-m unittest tests.test_guided_tool_dynamic_layout` -> 6 tests OK.
- Full tests before assemble: bundled Python `-m unittest discover -s tests` -> 26 tests OK.
- Rebuilt generated bundle with `python assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Full tests after assemble and practice bundle sync: bundled Python `-m unittest discover -s tests` -> 26 tests OK.
- Static VBA bundle validation: bundled Python `vbaeval.py validate` -> `OK 24 builders | balance, if-traps, wiring, strings, round-trip all clean`.
- No Office `~$` lock files were found in the workspace after the safe checks.

Sync status / blocker:
- Source files and generated text bundles are synced.
- Embedded `.docm` sync and Word runtime smoke were attempted next, but the Codex app escalation/usage gate rejected the Word COM automation request before it could run.
- Do **not** consider the two-column changes embedded in `CleanupSuite_Workspace.docm` or the practice `.docm` yet.
- Next safe step is to rerun the normal Word COM `.docm` sync and runtime smoke check when escalation/usage is available again.

## Session: 2026-06-14 00:20
Agent: Codex

Goal:
- Finish the previously blocked `.docm` sync and runtime verification for the two-column guided layout work.

Summary:
- Synced updated source VBA into both embedded macro documents through Word COM:
  - `CleanupSuite_Workspace.docm`,
  - `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`.
- Rebuilt `VBA_Cleanup_tool.txt` and refreshed `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Ran temporary smoke macros inside both `.docm` files and removed the temporary smoke modules afterward.

Runtime smoke coverage:
- Normal two-column choices:
  - `frmSoftReturnConverter`: `optSoftToPara` / `optParaToSoft`,
  - `frmFontNormalizer`: `chkFontFace` / `chkFontSize`.
- Custom-revealed two-column choices:
  - `frmUnicodeCleanup`: `chkNBSP` / `chkZWSP`,
  - `frmPunctuationCleanup`: `chkCurlyDouble` / `chkCurlySingle`.
- Conditional two-column choices:
  - `frmBreakNormalizer`: conversion choices hidden until conversion is enabled, then paired,
  - `frmDuplicateDetector`: fuzzy thresholds hidden until Fuzzy is selected, then paired,
  - `frmHeaderFooterStandardizer`: alignment choices hidden until Set alignment is enabled, then paired.

Tests/checks run:
- `python assemble.py` -> build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Bundled Python `-m unittest discover -s tests` -> 26 tests OK.
- Bundled Python `vbaeval.py validate` -> `OK 24 builders | balance, if-traps, wiring, strings, round-trip all clean`.
- Workspace `.docm` two-column smoke -> `OK`.
- Practice `.docm` two-column smoke -> `OK`.

Status:
- The earlier `.docm` sync blocker is resolved.
- Source, generated bundles, and both `.docm` files are now synced for the two-column guided layout pass.

## Session: 2026-06-14 11:00
Agent: Codex

Goal:
- Examine the current 0.7.x work, fix release-blocking issues if found, and push `0.7.4` as the code-freeze / final programming-side release candidate.

Backup created:
- `backups\CleanupSuite_backup_2026-06-14_10-59-32_v0.7.4-code-freeze-prep.zip`

Summary of release prep:
- Bumped current project metadata from `0.7.0` to `0.7.4` in:
  - `README.md`,
  - `VERSIONING.md`,
  - `src\modules\modCleanupLauncher.bas`,
  - the version assertion in `tests\test_launcher_redesign.py`.
- Rebuilt `VBA_Cleanup_tool.txt` and refreshed `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Synced updated source VBA into both embedded macro documents:
  - `CleanupSuite_Workspace.docm`,
  - `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`.
- Normalized newly changed blank `SEP` manifest lines so the release diff passes whitespace checks.

Exam / verification:
- `python assemble.py` -> build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Bundled Python `-m unittest discover -s tests` -> 26 tests OK.
- Bundled Python `vbaeval.py validate` -> `OK 24 builders | balance, if-traps, wiring, strings, round-trip all clean`.
- `git diff --check` -> clean.
- Embedded version read-back confirmed both `.docm` files contain `SUITE_VERSION = "0.7.4"`.
- Runtime layout smoke inside both `.docm` files returned `OK` for normal two-column choices, Custom-revealed choices, break conversion choices, duplicate fuzzy thresholds, and Header/Footer alignment choices.
- Preview OFF message remains covered by `tests\test_preview_action_panel.py`; a temporary COM smoke probe for that private click handler was removed because VBA cannot invoke the private handler from a temporary module.
- No Word processes or Office `~$` lock files remained after verification.

Code-freeze rule:
- Treat this `0.7.4` state as the programming-side code-freeze candidate for the `0.7.x` line.
- After `0.7.4` is pushed, do not make feature/code changes for `0.7.5` unless required to correct documentation-blocking bugs.

## Session: 2026-06-15 03:25
Agent: Codex

Goal:
- Add ChatGPT's `cleanup_suite_docs_0_7_4_addon.zip` documentation package to the locked local 0.7.4 release as documentation-only work.

Backup created:
- `backups\CleanupSuite_backup_2026-06-15_03-25-35_before_0.7.4_docs_addon.zip`

Files added:
- `docs\Human_Friendly_User_Manual_v0.7.4.md`
- `docs\Tool_Reference_Refined_v0.7.4.md`
- `docs\Programmers_Guide_Adding_Tools_v0.7.4.md`
- `docs\Documentation_Addon_Handoff_v0.7.4.md`
- `docs\README_DOC_PACKAGE_0.7.4.md`
- `docs\CODEX_ADD_0.7.4_DOCS_TASK.md`
- `documents\CleanupSuite_Human_Friendly_User_Manual_v0.7.4.docx`
- `documents\CleanupSuite_Programmers_Guide_Adding_Tools_v0.7.4.docx`

Documentation wording reviewed against locked 0.7.4:
- Updated the user manual global settings wording to match the actual launcher:
  - `Show completion review after apply`,
  - `Return to main menu after completion review`,
  - `Return to main menu after closing individual tool menus`,
  - `Auto-save before running each tool`.
- Updated the documentation basis notes to say the docs were reviewed against the local locked 0.7.4 project before commit.
- Updated developer-guide launcher wiring from the old direct `Me.Hide` / `.Show` pattern to the locked 0.7.4 `OpenCleanupTool "frm..."` pattern, with a note about fresh `VBA.UserForms.Add(...)` instances preventing the MSForms "form already displayed" runtime error.
- Updated tool headings to match actual 0.7.4 form titles:
  - `Direct Formatting Stripper` -> `Formatting Stripper`,
  - `Remove Objects & Elements` -> `Object Remover`.
- Mirrored relevant wording updates in the DOCX files where those sections existed.

Notes:
- No screenshots were added.
- No code behavior changes were made.

## Session: 2026-06-13 22:20
Agent: Codex

Goal:
- Finish the 0.7.0 guided submenu pass by applying the Capitalization-style functional design rules to the remaining tool menus, not just the visual shell.

Backup created:
- `backups\CleanupSuite_backup_2026-06-13_21-52-35_before-all-tool-design-pass.zip`

Summary of changes:
- Updated the shared `LayoutCleanupToolForm` runtime helper so generic tool menus are state-aware:
  - hides stale duplicate title labels by caption as well as by legacy title-label names,
  - hides custom option groups until `Custom` is selected,
  - hides break conversion choices until `Convert section breaks` is enabled,
  - hides fuzzy duplicate thresholds until `Fuzzy match` is selected,
  - hides Header/Footer alignment choices until `Set alignment` is selected,
  - hides Header/Footer standardize-only controls when `Clear all header/footer content` is selected,
  - hides blank `optScopeSelection` controls so whole-document-only tools do not show an empty radio button.
- Added relayout hooks to custom and conditional forms so choosing modes immediately redraws the menu:
  - Unicode, Punctuation, Spacing, Lists, Paragraphs, Capitalization,
  - Break Normalizer,
  - Duplicate Paragraph Detector,
  - Header / Footer Standardizer.
- Fixed Select All / Deselect All guards to use `optCustom.Value` instead of hidden frame visibility.
- Swapped the first two Main Menu global settings so `Show completion review after apply` appears before `Return to main menu after completion review`.
- Rebuilt generated bundles and synced both `.docm` files.

Tests/checks run:
- Focused RED test first failed against the old behavior as expected.
- Focused test after implementation: bundled Python `-m unittest tests.test_guided_tool_dynamic_layout` -> 4 tests OK.
- Full tests before assemble: bundled Python `-m unittest discover -s tests` -> 24 tests OK.
- Rebuilt generated bundle with bundled Python `assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Synced `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Synced both embedded `.docm` files through Word COM with macros disabled while opening:
  - `CleanupSuite_Workspace.docm`,
  - `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`.
- Embedded workspace `.docm` smoke macro returned `OK` for:
  - Unicode custom controls hidden by default and visible after Custom,
  - Break conversion choices hidden by default and visible after enabling conversion,
  - Duplicate fuzzy threshold controls hidden by default and visible after Fuzzy,
  - Header/Footer alignment and clear-all state visibility,
  - Document Trim blank scope radio hidden,
  - Punctuation has exactly one visible title label.
- Final full test discovery: bundled Python `-m unittest discover -s tests` -> 24 tests OK.

## Session: 2026-06-15 06:00
Agent: Codex

Goal:
- Create the `0.7.4.5` bugfix bridge before the documentation-complete `0.7.5` release.
- Fix release-blocking issues without changing tool feature behavior:
  - Preview Actions summary clipping for multi-line summaries.
  - Spacing Fixer compile error.
  - Table Cleaner compile error.

Backup created:
- `backups\CleanupSuite_backup_2026-06-15_06-02-12_v0.7.4.5_bugfix_pre_sync.zip`

Summary of changes:
- Bumped release metadata from `0.7.5` to `0.7.4.5` in source/tests/README/VERSIONING so Word reports the bridge bugfix version.
- Increased the shared Preview Actions panel height from `90` to `100` and the summary label height from `16` to `26`, in both source and installer-generated layout.
- Fixed Spacing Fixer duplicate local declaration by renaming the apply-path punctuation array to `applyPunctPatterns`.
- Fixed Table Cleaner compile error by replacing invalid `tbl.Range.ClearFormatting` with Word-supported `tbl.Range.Font.Reset` and `tbl.Range.ParagraphFormat.Reset`.
- Added regression tests for Preview Actions dimensions and the two compile-prone source patterns.
- Rebuilt `VBA_Cleanup_tool.txt`, synced the practice copy, and synced both `.docm` files through Word COM before later smoke attempts.

Tests/checks run:
- Focused RED checks first failed against the old Preview Actions height, Spacing duplicate declaration, and Table Cleaner invalid range formatting call.
- Focused tests after implementation: `python -m unittest tests.test_preview_action_panel tests.test_tool_compile_regressions tests.test_launcher_redesign` -> 19 tests OK.
- Rebuilt generated bundle with `python assemble.py`; build-gate validators passed (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Full test discovery: `python -m unittest discover -s tests` -> 28 tests OK.
- Validator: `python vbaeval.py validate` -> OK.
- `git diff --check` -> OK.
- Source/generated checks confirmed:
  - `Const FORM_H As Single = 100`,
  - `lblSummary.Move M, 40, CONTENT_W, 26`,
  - generated bundles carry `SUITE_VERSION = "0.7.4.5"`,
  - no remaining `tbl.Range.ClearFormatting`,
  - no duplicate `punctPatterns` declaration pattern.

Notes:
- Word COM runtime smoke/readback attempts hung after hidden modal automation prompts. Headless `WINWORD` processes and Office lock stubs were cleaned up.
- Because the earlier `.docm` sync command completed successfully before the smoke attempts, both `.docm` files should contain the updated code, but a later embedded readback could not be completed due to Word COM hanging.

## Session: 2026-06-15 11:31
Agent: Codex

Goal:
- Finalize the documentation-complete `0.7.5` release from the pushed `0.7.4.5` bugfix bridge state.
- Add/update the 0.7.5 documentation wording to reflect that `0.7.5` includes the locked `0.7.4` work plus the `0.7.4.5` release-blocking bridge fixes.

Backup created:
- `backups\CleanupSuite_backup_2026-06-15_11-31-37_before-v0.7.5-final-from-0.7.4.5.zip`

Summary of changes:
- The Google Drive package link required sign-in and could not be downloaded directly in automation; the failed placeholder download was removed.
- Used the existing local `0.7.5` Markdown/DOCX documentation package and adjusted Markdown wording against the actual locked local `0.7.4.5` behavior.
- Updated project version metadata from `0.7.4.5` to `0.7.5`.
- Rebuilt generated bundles with `python assemble.py`.
- Synced `VBA_Cleanup_tool.txt` to `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Synced both `.docm` files through Word COM using temp copies, then replaced the workspace and practice documents.
- No feature/UI/tool behavior changes were made beyond carrying forward the already-pushed `0.7.4.5` bridge fixes and the required `0.7.5` version metadata.

Documentation notes:
- Updated the 0.7.5 user manual, tool reference, programmer guide, release notes, and documentation handoff wording so they honestly describe the `0.7.4.5` bridge.
- Added the Word COM caveat to release documentation, while also recording that this session completed embedded `.docm` readback successfully.
- DOCX files were inspected; they did not contain the detailed basis/caveat wording that needed Markdown correction, so they were left unchanged.

Tests/checks run:
- `python assemble.py` -> OK (`24 builders | balance, if-traps, wiring, strings, round-trip all clean`).
- Embedded `.docm` readback on temp copies -> OK for both workspace documents; verified `SUITE_VERSION = "0.7.5"`, Preview Actions height/summary dimensions, Spacing compile fix, and Table Cleaner compile fix.
- `python -m unittest discover -s tests` -> 28 tests OK.
- `python vbaeval.py validate` -> OK.
- `git diff --check` -> OK.
- No lingering `WINWORD` process, temp sync file, or Office lock stub remained after sync/readback.
