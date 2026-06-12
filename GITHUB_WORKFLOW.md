# GitHub Workflow — CleanupSuite for MS Word

This project is also in a GitHub repo. Chris wants Codex and ChatGPT to be able to work as careful coworkers.

```text
Repo: MasseysLab/CleanupSuiteForWord
URL: https://github.com/MasseysLab/CleanupSuiteForWord
Default branch observed: main
```

## General rule

Use Git history as one layer of project protection, but do not treat Git as a replacement for local backup zips. For local Word/VBA projects, especially with `.docm` files, keep both Git commits and timestamped backup zips.

## Recommended workflow

1. Start each local session by reading:
   - `AI_HANDOFF.md`
   - `AI_RULES.md`
   - `REPO_FINDINGS_AND_SYNC_CHECK.md` if working on Preview Actions or sync issues
   - any specific task file named in the latest handoff, such as `PREVIEW_ACTIONS_NEXT_TASK.md`.

2. Create a local backup zip before editing.

3. Check Git status:

```powershell
git status
```

4. Prefer a branch for meaningful changes:

```powershell
git checkout -b ai/session-YYYY-MM-DD-short-description
```

5. Keep commits focused:
   - Safety setup commit: handoff/rules/scripts/docs.
   - Code behavior commits: actual project changes.
   - Test commits: test updates, if separate.

6. Before reporting success, check:

```powershell
git status
git diff --stat
```

7. Append to `AI_HANDOFF.md` with:
   - branch name
   - files changed
   - tests/checks run
   - known issues
   - next safest step

## Important sync warning

GitHub `main` may not reflect Chris's local working folder or `.docm`. Do not discard uncommitted local changes. Before using GitHub as the source of truth, compare it to the local files and runtime `.docm` behavior.

## ChatGPT collaboration rule

ChatGPT may be asked to work through GitHub or uploaded files. ChatGPT can read, create, update, and delete files when Chris explicitly asks for project work, but should:

- keep changes small;
- avoid unnecessary deletes;
- prefer branch/PR-style work when possible;
- update `AI_HANDOFF.md` after making changes;
- be clear about anything it could not verify locally.

## Codex local advantage

Codex may be better for local `.docm` and Word/VBA operations because it can work inside Chris's local folder and use local tooling. ChatGPT is usually better suited for text files, reviews, diffs, explanations, instructions, tests, and GitHub/source-level changes.
