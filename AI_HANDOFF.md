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

## Current release snapshot

- Current working release in source: `0.8.5` Alpha
- Current active working branch: `codex-0.8.0-release`
- Current upstream branch: `origin/codex-0.8.0-release`
- Legacy remote branch `origin/codex-0.7.0-start` may still exist as older history and can be cleaned up later if Chris wants.

## Current release documentation set

Current `docs/` files:

- `docs/Human_Friendly_User_Manual_v0.8.5.md`
- `docs/Programmers_Guide_Adding_Tools_v0.8.5.md`
- `docs/Tool_Reference_Refined_v0.8.5.md`
- `docs/Release_Notes_0.8.0_to_0.8.5.md`
- `docs/Documentation_Handoff_0.8.5.md`
- `docs/README_DOC_PACKAGE.md`

Current `documents/` release artifacts:

- `documents/CleanupSuite_Human_Friendly_User_Manual_v0.8.0.pdf`
- `documents/CleanupSuite_Human_Friendly_User_Manual_v0.8.0.docx`
- `documents/CleanupSuite_Programmers_Guide_Adding_Tools_v0.8.0.docx`
- `documents/CleanupSuite_Tool_Reference_Refined_v0.8.0.docx`

## Current tool-form state

- Guided tool forms use the shared layout path in `LayoutCleanupToolForm`.
- Visible tool titles belong in the form body, not the native title bar.
- Help stays on the main menu, not inside tool forms.
- Header / Footer Standardizer uses the special row-order rule captured in `docs/Tool-Form Rules.md`.
- When changing guided-form layout behavior, read `docs/Tool-Form Rules.md` before editing.

## Current workflow reminders

1. Read this file first.
2. Read `docs/Tool-Form Rules.md` before touching tool-form layout behavior.
3. Create a fresh backup before risky edits or Word/COM sync work.
4. Do not hand-edit `VBA_Cleanup_tool.txt`; edit `src/` and rebuild with `python assemble.py`.
5. If generated root artifacts change, update the matching practice artifacts in `Practice - Try CleanupSuite Here`.
6. If release metadata changes, sync source/generated/docm state and update the current-facing docs.

## Recent high-value milestones

- `0.8.0` source/version wording was cleaned up across current-facing release files.
- The minimal `0.8.0` documentation package was expanded into a credible current doc set.
- `0.8.0` DOCX export artifacts were generated for the user manual, programmer guide, and tool reference.
- The working branch was renamed from `codex-0.7.0-start` to `codex-0.8.0-release`.
- This handoff was trimmed to current state, and older long-form session history was moved to `AI_HANDOFF_ARCHIVE.md`.
- `docs/Tool-Form Rules.md` and `AI_RULES.md` were updated to reflect the current course, compact handoff pattern, release-doc package expectations, and current branch identity.
- `0.8.5` Alpha: guided info box redesigned from single-line to two-column layout (display name + advisory text); `GuidedInfoDisplayName()` and `GuidedInfoAdvisory()` added for all current controls; "Custom (choose below)" → "Custom" across 5 forms; preview form title blanked; option heights and long-caption widths adjusted.
- `0.8.5` version bump: `SUITE_VERSION` updated to `"0.8.5"`, versioned docs created, release notes written.
- Final closeout checks passed: `python assemble.py`, `python vbaeval.py validate`, `python -m unittest discover -s tests`, and `git diff --check`.

## Next milestone

- Word smoke test per `docs/Alpha_Readiness_Checklist_0.8.5.md` — the code and docs are ready; this is the remaining gate before calling `0.8.5` fully done.
- Keep `codex-0.8.0-release` as the working branch; do not merge to `main` casually because `main` is still far behind.
- Do not clean house before `0.9.0` documentation is complete unless Chris explicitly asks.

## Session: 2026-06-24 22:18
Agent: Codex

Goal:
- Refresh `CleanupSuite_Workspace.docm` and `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm` from current local source.

Backup created:
- `backups/CleanupSuite_backup_2026-06-24_22-14-17.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.8.0-release`
- Status before: dirty working tree with existing `0.8.5` Alpha source/doc/generated/docm changes already in progress.
- Status after: same broader dirty working tree, plus `scripts/sync_docm_code_only.ps1` and `tests/test_docm_sync_script.py` updated/added to keep future `.docm` syncs complete.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root/practice bundles, and both local `.docm` files.

Files changed:
- `scripts/sync_docm_code_only.ps1`
- `tests/test_docm_sync_script.py`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`

Summary of changes:
- Rebuilt `VBA_Cleanup_tool.txt`, which also refreshed `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`.
- Found that the docm sync script skipped `modCleanupLauncher`, `frmPreviewActions`, and `frmCleanupSuiteLauncher`, which left embedded docs stale even after a reported sync.
- Added those components to the sync script and resynced both `.docm` files.

Important observations:
- Before the script fix, embedded read-back showed stale `SUITE_VERSION = "0.8.0"` in `modCleanupLauncher` and stale `frmPreviewActions` code while source was already `0.8.5`.
- After the fix, read-back confirmed both `.docm` files match current source for `modCleanupLauncher`, `frmPreviewActions`, `frmCleanupSuiteLauncher`, `modCleanupHelpers`, and `modCleanupSuiteInstaller`.

Tests/checks run:
- `python assemble.py`
- `powershell -ExecutionPolicy Bypass -File scripts\sync_docm_code_only.ps1`
- `python -m unittest tests.test_docm_sync_script` (red, then green after script fix)
- `python vbaeval.py validate`
- `python -m unittest discover -s tests`
- Word COM read-back comparison against source for both `.docm` files

Known remaining issues:
- Full manual Word smoke testing is still the next functional gate for `0.8.5` Alpha.

Next recommended step:
- Run the Word smoke test checklist in `docs/Alpha_Readiness_Checklist_0.8.5.md` against the refreshed practice `.docm`.

## Session: 2026-06-25 00:05
Agent: Codex

Goal:
- Replace the Capitalization Fixer's Word-duplicate case modes with a deterministic smart-repair tool.

Backup created:
- `backups/CleanupSuite_backup_2026-06-24_23-27-20.zip`

Git/repo state:
- Remote: `origin https://github.com/MasseysLab/CleanupSuiteForWord.git`
- Branch: `codex-0.8.0-release`
- Status before: dirty working tree with existing `0.8.5` Alpha source/doc/generated/docm changes already in progress.
- Status after: same broader dirty working tree, plus the smart-capitalization repair source/tests/docs/generate/docm updates from this session.
- GitHub/local/generated/.docm sync checked: yes for local source, generated root/practice bundles, and the root `.docm`; the practice `.docm` sync command succeeded after the file lock was cleared, but a separate final read-only Word COM verification pass on that file timed out before returning component comparisons.

Files changed:
- `src/forms/frmCapitalizationCleanup.bas`
- `src/modules/modCleanupHelpers.bas`
- `src/installer/installer.bas`
- `tests/test_capitalization_guided_form.py`
- `tests/test_guided_tool_dynamic_layout.py`
- `tests/test_shared_help_and_manual.py`
- `docs/Capitalization_Smart_Repair_Followups.md`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here\VBA_Cleanup_tool.txt`
- `CleanupSuite_Workspace.docm`
- `Practice - Try CleanupSuite Here\CleanupSuite_Workspace.docm`

Summary of changes:
- Reworked Capitalization Fixer into a three-mode deterministic repair tool with `Conservative`, `Balanced`, and `Aggressive` modes.
- Replaced Word-duplicate casing choices with advanced protection toggles for abbreviation lists, acronym protection, name/brand protection, heading heuristics, and quote/bullet/boundary context.
- Added a minimal built-in data architecture using compiled abbreviation, acronym, and protected-term lookup dictionaries plus deterministic paragraph and sentence heuristics.
- Recorded future-version follow-up work in `docs/Capitalization_Smart_Repair_Followups.md`.
- Rebuilt generated bundles and synced both workspace `.docm` files after the practice-file lock was cleared.

Important observations:
- VBA source must stay ASCII-safe when it will be synced through Word/VBProject code modules; Unicode dash/bullet literals inside the form code had to be converted to `ChrW$` references to keep embedded read-back stable.
- The existing `scripts/sync_docm_code_only.ps1` path worked for both docs after the practice-file lock was cleared by the user.

Tests/checks run:
- `python assemble.py`
- `python vbaeval.py validate`
- `python -m unittest discover -s tests`
- `git diff --check`
- `powershell -ExecutionPolicy Bypass -File scripts\sync_docm_code_only.ps1`
- Embedded read-back comparison on `CleanupSuite_Workspace.docm` for `modCleanupLauncher`, `frmCapitalizationCleanup`, `modCleanupHelpers`, and `modCleanupSuiteInstaller` -> all matched source.

Known remaining issues:
- Run a manual Word smoke test on the refreshed capitalization tool, especially around headings, bullets, abbreviations, and long ALL-CAPS words.
- If needed, perform one more separate practice `.docm` read-back after Word is fully idle, since the final read-only verification pass on that file timed out even though the sync command itself succeeded.

Next recommended step:
- Open the practice `.docm` and smoke-test the new `Conservative` / `Balanced` / `Aggressive` modes against realistic messy text before committing.

## Archive note

The previous long-form handoff history was preserved in:

```text
AI_HANDOFF_ARCHIVE.md
```

Use the archive when older release history, old backup names, or previous session-by-session details are needed.
