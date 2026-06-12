# Codex Setup Task — CleanupSuite Collaboration Safety

Chris wants Codex to prepare the project so Codex and ChatGPT can safely alternate as coworkers on the same project.

Project folder:

```text
C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word
```

GitHub repo:

```text
https://github.com/MasseysLab/CleanupSuiteForWord
MasseysLab/CleanupSuiteForWord
```

## Critical setup rule

Create a backup before copying these files into the project root. If any of these files already exist, do not blindly overwrite valuable handoff history. Merge or append carefully, and record what you did.

## Codex task list

1. Copy these files into the project root:
   - `AI_HANDOFF.md`
   - `AI_RULES.md`
   - `CODEX_SETUP_TASK.md`
   - `GITHUB_WORKFLOW.md`
   - `REPO_FINDINGS_AND_SYNC_CHECK.md`
   - `PREVIEW_ACTIONS_NEXT_TASK.md`
   - `Start-AI-Session.ps1`
   - `README_FOR_CHRIS.txt`

2. Before changing anything else, create a local timestamped backup zip of the full project folder.
   - Put it in `backups\`.
   - Do not include the `backups` folder inside the zip.
   - Do not overwrite older backups.

3. Make sure backups are not committed to Git.
   - If the repo has a `.gitignore`, add `backups/` if missing.
   - If there is no `.gitignore`, create one only if appropriate, and include `backups/`.

4. Check the Git remote and branch.
   - Record repo URL and branch in `AI_HANDOFF.md`.
   - Prefer branch/PR workflow for non-trivial changes.

5. Confirm these files are present and readable:
   - `AI_HANDOFF.md`
   - `AI_RULES.md`
   - `CODEX_SETUP_TASK.md`
   - `GITHUB_WORKFLOW.md`
   - `REPO_FINDINGS_AND_SYNC_CHECK.md`
   - `PREVIEW_ACTIONS_NEXT_TASK.md`
   - `Start-AI-Session.ps1`

6. Double-check the setup before saying it is complete.

7. Append a new Codex entry to `AI_HANDOFF.md` after setup is complete.

## Required double-check list

Codex should explicitly verify and record:

- Backup zip created: yes/no, path recorded.
- Backup zip excludes `backups/`: yes/no.
- `backups/` ignored by Git: yes/no.
- Required files copied into root: yes/no.
- Required files readable: yes/no.
- Git remote recorded: yes/no.
- Current branch recorded: yes/no.
- Working tree status before/after recorded: yes/no.
- If existing handoff files existed, their contents were preserved/merged: yes/no/not applicable.
- Handoff entry appended: yes/no.

If any item is no or unknown, record why.

## Additional repo-local sync check

Before working on Preview Actions, read `REPO_FINDINGS_AND_SYNC_CHECK.md`. GitHub `main` may not match Chris's local UI state.

Run and record a summary:

```powershell
git remote -v
git branch --show-current
git status --short
git log --oneline -5
git diff -- src/forms/frmPreviewActions.bas src/installer/installer.bas tests/test_preview_action_panel.py VBA_Cleanup_tool.txt
git grep -n -e "Preview Actions" -e "frmPreviewActions" -e "lblTitle" -e "cmdPreview" -e "cmdReconfigure" -e "cmdApply" -e "cmdClear"
```

## Suggested Codex handoff entry after setup

```markdown
## Session: YYYY-MM-DD HH:mm
Agent: Codex

Goal:
- Install AI collaboration safety files and prepare backup/Git workflow.

Backup created:
- backups\CleanupSuite_backup_YYYY-MM-DD_HH-mm-ss.zip

Git/repo state:
- Remote: <fill in>
- Branch: <fill in>
- Working tree status before setup: <fill in>
- Working tree status after setup: <fill in>

Files changed:
- AI_HANDOFF.md
- AI_RULES.md
- CODEX_SETUP_TASK.md
- GITHUB_WORKFLOW.md
- REPO_FINDINGS_AND_SYNC_CHECK.md
- PREVIEW_ACTIONS_NEXT_TASK.md
- Start-AI-Session.ps1
- README_FOR_CHRIS.txt
- .gitignore, if updated

Summary of changes:
- Added project safety rules and handoff process for Codex/ChatGPT collaboration.
- Added backup script and workflow instructions.
- Confirmed or added Git ignore rule for local backup zips.

Important observations:
- Repo remote: <fill in>
- Current branch: <fill in>
- Working tree status before setup: <fill in>
- Working tree status after setup: <fill in>

Tests/checks run:
- Confirmed backup zip was created.
- Confirmed backup zip does not include the backups folder.
- Confirmed handoff/rules/repo findings files are readable.
- Confirmed Git status.

Known remaining issues:
- Preview Actions duplicate text remains the next known technical task unless already fixed.
- GitHub/local/generated/.docm sync must be checked before fixing Preview Actions.

Next recommended step:
- Continue with `REPO_FINDINGS_AND_SYNC_CHECK.md` and `PREVIEW_ACTIONS_NEXT_TASK.md` only after safety setup is verified.
```
