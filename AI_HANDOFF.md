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

- Current working release in source: `0.8.0`
- Current active working branch: `codex-0.8.0-release`
- Current upstream branch: `origin/codex-0.8.0-release`
- Legacy remote branch `origin/codex-0.7.0-start` may still exist as older history and can be cleaned up later if Chris wants.

## Current release documentation set

Current `docs/` files:

- `docs/Human_Friendly_User_Manual_v0.8.0.md`
- `docs/Programmers_Guide_Adding_Tools_v0.8.0.md`
- `docs/Tool_Reference_Refined_v0.8.0.md`
- `docs/Release_Notes_0.7.5_to_0.8.0.md`
- `docs/Documentation_Handoff_0.8.0.md`
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
- `docs/Tool-Form Rules.md` and `AI_RULES.md` were updated to reflect the current `0.8.0` course, compact handoff pattern, release-doc package expectations, and current branch identity.
- Final closeout checks passed: `python assemble.py`, `python vbaeval.py validate`, `python -m unittest discover -s tests`, and `git diff --check`.
- Roadmap clarified: `0.8.5` is Alpha, `0.9.0` is Alpha plus full documentation, and clean-house work waits until that `0.9.0` documentation state is complete unless Chris explicitly asks earlier.

## Next milestone

- Use `docs/Alpha_Readiness_Checklist_0.8.5.md` for the next focused milestone.
- Keep `codex-0.8.0-release` as the preserved pushed release branch for now; do not merge to `main` casually because `main` is still far behind.
- Do not clean house before `0.9.0` documentation is complete unless Chris explicitly asks.

## Archive note

The previous long-form handoff history was preserved in:

```text
AI_HANDOFF_ARCHIVE.md
```

Use the archive when older release history, old backup names, or previous session-by-session details are needed.
