CleanupSuite AI Collaboration Setup Package — v3

Put the contents of files_to_copy_to_project_root directly into:

C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite for MS Word

Then tell Codex:

Before editing anything, read AI_HANDOFF.md, AI_RULES.md, CODEX_SETUP_TASK.md, and REPO_FINDINGS_AND_SYNC_CHECK.md. Create a backup zip first, verify the backup and Git ignore rules, then append a setup entry to AI_HANDOFF.md. The next technical task is documented in PREVIEW_ACTIONS_NEXT_TASK.md.

Included files:
- Codex_Project_Safety_and_Handoff_Setup_v3.docx
- AI_HANDOFF.md
- AI_RULES.md
- CODEX_SETUP_TASK.md
- GITHUB_WORKFLOW.md
- REPO_FINDINGS_AND_SYNC_CHECK.md
- PREVIEW_ACTIONS_NEXT_TASK.md
- Start-AI-Session.ps1
- README_FOR_CHRIS.txt

Critical v3 addition:
GitHub main may not match Chris's local working folder or the actual .docm. Codex must compare repo/local/generated/.docm state before fixing Preview Actions.

ChatGPT could not create the first real local backup because it cannot directly access the Windows project folder. Codex should create that backup locally before editing.
