# CleanupSuite Release Notes - 0.7.4 to 0.7.5

**Release:** 0.7.5
**Date:** 2026-06-15

## Meaning

CleanupSuite 0.7.5 is the documentation-complete release for the 0.7.x line.

Version 0.7.4 remains the programming-side code-freeze checkpoint. Version 0.7.4.5 was a small bridge release for release-blocking Preview Actions and compile fixes. Version 0.7.5 uses that corrected behavior and adds finalized user/developer documentation plus release identity updates.

## Added

- Human-friendly user manual for CleanupSuite 0.7.5.
- Programmer's guide for adding tools.
- Refined tool reference with more specific preview and safety guidance.
- Documentation handoff note for the 0.7.5 release.
- DOCX versions of the user manual and programmer's guide.

## Changed

- Updated project release wording from 0.7.4 code checkpoint to 0.7.5 documented release.
- Updated `SUITE_VERSION` to `0.7.5` so Word identifies this documented release correctly.
- Rebuilt generated bundles and synced the working/practice macro documents after the version metadata change.
- Carried forward the 0.7.4.5 bridge fixes: taller Preview Actions summaries, Spacing Fixer compile repair, and Table Cleaner compile repair.

## Not Changed

- No feature behavior changes beyond the 0.7.4.5 release-blocking fixes.
- No screenshots were added in this release; screenshots are still planned for a later documentation pass.

## Verification Note

The 0.7.4.5 bridge sync completed before later Word COM smoke/readback attempts hung on hidden modal prompts. For 0.7.5, Codex should complete the embedded `.docm` readback if Word COM cooperates; otherwise the release handoff should record the limitation plainly.
