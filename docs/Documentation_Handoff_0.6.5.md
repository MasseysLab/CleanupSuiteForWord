# CleanupSuite Documentation Handoff — 0.6.5

**Prepared:** 2026-06-12  
**Codebase documented:** 0.6.4 code-freeze candidate  
**Commit documented:** `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`

## Intent

Add documentation only, then release 0.6.5. The 0.6.5 release should use the same code as 0.6.4 unless a documentation-blocking bug is discovered.

## Files to add

- `docs/User_Manual_v0.6.5.md`
- `docs/Quick_Start_v0.6.5.md`
- `docs/Troubleshooting_v0.6.5.md`
- `docs/Release_Notes_0.6.4_to_0.6.5.md`
- `docs/Documentation_Handoff_0.6.5.md`
- `documents/CleanupSuite_User_Manual_v0.6.5.docx` if the project stores distributable docs in-repo

## Codex checklist

1. Create a fresh backup before touching the repo.
2. Confirm the working tree is on top of the pushed 0.6.4 commit.
3. Add the documentation files.
4. Do not change source code unless a documentation-blocking bug is found and Chris approves.
5. Update README/VERSIONING only as needed to mark 0.6.5 as the documentation-complete release.
6. Run normal tests if any source/version files changed.
7. Append an `AI_HANDOFF.md` entry explaining that 0.6.5 is documentation-only relative to 0.6.4.
8. Commit with a message such as `Release 0.6.5 documentation complete`.
9. Push to GitHub.
10. Do not start 0.7.0 work until 0.6.5 is pushed.

## Suggested version wording

```text
0.6.4 = final programming-side code freeze
0.6.5 = documentation-complete release using the same codebase as 0.6.4
0.7.0 = next development milestone
```
