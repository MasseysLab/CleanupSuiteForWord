# AI Rules — CleanupSuite for MS Word

These rules apply to both ChatGPT and Codex when working on this project.

## Permission model

Chris has approved Codex to change, add, and delete files within the local project folder.

Chris also wants ChatGPT to have similar permission when working through the GitHub repo or supplied project files. Because ChatGPT may not have direct access to the local Windows folder, ChatGPT changes may be delivered as GitHub commits/branches, downloadable files, zips, or patch-style instructions for Codex or Chris to apply locally.

Both agents should act as careful coworkers, not as independent owners of the project.

## Repo information

```text
GitHub repo: https://github.com/MasseysLab/CleanupSuiteForWord
Repository full name: MasseysLab/CleanupSuiteForWord
Default branch observed: main
Local project folder: C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word
```

## Mandatory startup steps

1. Read `AI_HANDOFF.md`.
2. Read this file, `AI_RULES.md`.
3. Create a timestamped backup zip before editing anything locally.
4. Inspect the current project state before changing files.
5. Check Git remote, current branch, and status before editing when working locally.
6. If working on Preview Actions or a `.docm`/runtime bug, read `REPO_FINDINGS_AND_SYNC_CHECK.md` and compare GitHub/local/generated/embedded state.
7. If a previous handoff says something is unresolved, read that section before doing unrelated work.

## Do not destroy local work

Before any reset, clean, branch switch, pull, merge, rebase, generated-file overwrite, or `.docm` replacement:

- create a backup zip;
- record Git status;
- inspect local diffs;
- preserve or copy out any uncommitted work;
- ask Chris if it is unclear which version is correct.

## Mandatory verification before saying "done"

Codex or ChatGPT must verify the work, not just make the change.

For setup changes, verify:

- required files exist;
- handoff/rules/repo findings files are readable;
- backup zip exists;
- backup zip does not contain the `backups` folder;
- `backups/` is ignored by Git;
- Git branch, remote, and status are recorded.

For code/UI changes, verify as appropriate:

- source files changed as intended;
- generated `VBA_Cleanup_tool.txt` was updated if source changed;
- embedded `.docm` was updated if the `.docm` is part of the deliverable;
- validators/tests were run when available;
- manual runtime check was performed when the bug is visual or Word/VBA-specific;
- GitHub/local/generated/.docm sync was checked when relevant;
- `AI_HANDOFF.md` was updated with exact files, checks, and remaining issues.

If something cannot be verified, say so clearly and record it in `AI_HANDOFF.md`.

## Mandatory shutdown steps

After making changes:

1. Append a new entry to `AI_HANDOFF.md`.
2. List every file changed.
3. Explain what changed and why.
4. Record tests/checks run.
5. Record anything not fixed or not attempted.
6. Record the safest next step.

## Backup rule

Create backups in:

```text
backups```

Use this naming format:

```text
CleanupSuite_backup_YYYY-MM-DD_HH-mm-ss.zip
```

Do not overwrite old backups.

The `backups` folder itself should not be included inside new backup zips.

The `backups/` folder should normally be ignored by Git.

## GitHub workflow

1. Prefer a branch or PR-style workflow for non-trivial changes.
2. Keep `main` or the default branch stable unless Chris explicitly asks for direct commits.
3. Commit handoff/rules/docs changes separately from code behavior changes when practical.
4. Use clear commit messages.
5. Do not delete files unless the reason is clear and recorded in `AI_HANDOFF.md`.
6. When possible, compare diffs before reporting success.

## Editing rules

1. Keep changes small and reversible.
2. Do not make broad rewrites unless Chris specifically asks for them.
3. Do not assume GitHub `main`, local source files, generated `VBA_Cleanup_tool.txt`, and embedded `.docm` VBA/forms are in sync.
4. If `.docm` output matters, confirm whether the actual working `.docm` was updated.
5. For UI bugs, audit actual runtime controls before changing form layout.
6. For Word/VBA form issues, check labels, frames, pages, multipages, form captions, and dynamically created controls.
7. Do not keep changing a control just because its name sounds relevant; prove it is the visible control first.
8. Preserve tests unless they are clearly wrong and explain any test changes.
9. If unsure, stop and explain instead of guessing.
10. Only one AI should work on the project at a time.

## Preferred handoff entry format

```markdown
## Session: YYYY-MM-DD HH:mm
Agent: Codex / ChatGPT

Goal:
- ...

Backup created:
- ...

Git/repo state:
- Remote:
- Branch:
- Status before:
- Status after:
- GitHub/local/generated/.docm sync checked:

Files changed:
- ...

Summary of changes:
- ...

Important observations:
- ...

Tests/checks run:
- ...

Known remaining issues:
- ...

Next recommended step:
- ...
```
