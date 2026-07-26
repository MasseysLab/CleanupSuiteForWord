# CleanupSuite Last Pre-Hybrid Task List

**Checkpoint date:** July 18, 2026
**Checkpoint meaning:** Last major persistent backup before hybrid-engine implementation begins
**Published release:** `0.9.2-alpha`
**Development branch:** `codex/0.9.5-beta`
**Architecture:** C# analysis sidecar with VBA as the Word mediator—approved but not implemented
**Authorization state:** All new implementation remains paused until Chris explicitly says `go`

**Release-track decision:** Chris confirmed that development moves forward as Beta only. The current local work is the `0.9.5 Beta` baseline; no additional Alpha hotfix will be prepared.

This task list preserves the ordered work plan at the last pre-hybrid checkpoint. “Pre-hybrid” means no C# engine code exists yet; it does not mean the architectural decision is undecided. The architecture is approved, while implementation, commits, tags, pushes, and publication remain separately gated.

```text
Accepted Beta baseline → Safety foundation → Hybrid foundation → Pilot tools → Beta features → Installer → Release gate
```

## 1. Preserve the Accepted Local Beta Baseline

**Status: Accepted as the Beta baseline; implemented locally but not committed or published**

- Confirm minimal one-character Preview markers.
- Confirm Duplicate Paragraph Remover behavior.
- Fold these changes into `0.9.5-beta`; do not prepare another Alpha hotfix.
- Preserve the punctuation fast path and stable Reading View behavior.

Estimated combined difficulty: **8–15%**.
Estimated professional desirability: **96–99%**.

## 2. Implement the Approved Safety Package

**Status: Approved, waiting for `go`**

- Create the shared blank-paragraph classifier.
- Protect table-cell, row-end, and final-document markers.
- Ensure Spacing Fixer never deletes table rows.
- Make Paragraph Fixer, Duplicate Remover, and Document Trim use the classifier.
- Make Table Cleaner the sole owner of table-row and column deletion.
- Keep structural options OFF and outside Recommended/All defaults.
- Detect objects, controls, fields, bookmarks, nested tables, merged structures, and other meaningful table content.
- Make Preview and Apply consume the same candidates.
- Revalidate candidates and delete bottom-to-top.
- Correct single-column table-to-text Preview/Apply parity.
- Reset risky Break Normalizer choices to OFF.
- Use count-specific structural confirmation and one Word Undo record.
- Add removable/protected/skipped blank counts.

Estimated combined difficulty: **58–70%**.
Estimated professional desirability: **97–99%**.

## 3. Finish the Preview Safety and Navigation Foundation

**Status: Navigation approved; highlight preservation strongly recommended but not formally approved as a complete package**

- Implement the exact navigation layout:

  `[Change] [Page] ◀ Previous | Next ▶ [Page] [Change]`

- Give all four buttons identical dimensions.
- Use pale gold for Change, pale blue for Page, and gray when disabled.
- Store changed-page anchors during Preview.
- Add a compact inspection path for nonhighlightable findings.
- Preserve pre-existing user highlights and shading exactly.
- Eliminate whole-document `wdNoHighlight` cleanup behavior.
- Restore Preview state across Apply, Reconfigure, Close, errors, and document changes.

Estimated combined difficulty: **55–70%**.
Estimated professional desirability: **98–100%**.

## 4. Define the Hybrid Contract

**Status: Architecture approved, waiting for `go`**

- Define the versioned C#/VBA request-and-response schema.
- Finalize the candidate record and fingerprint rules.
- Finalize the standard Apply-operation vocabulary.
- Define protected, skipped, confidence, and review-only states.
- Define capability discovery and version compatibility.
- Decide the initial secure job/result-file locations and cleanup lifecycle.
- Add contract and compatibility tests before migrating tools.

Estimated combined difficulty: **42–57%**.
Estimated professional desirability: **97–100%**.

## 5. Build the Isolated C# Sidecar Foundation

**Status: Approved direction; no C# code exists yet**

- Create the self-contained C# engine.
- Keep it outside Word's process.
- Implement restricted UTF-8 job/result communication initially.
- Add logging that excludes document content by default.
- Add cancellation, error reporting, and temporary-file cleanup.
- Ensure the engine only returns candidates and never edits the open document.
- Add automated engine and protocol tests.

Estimated combined difficulty: **48–62%**.
Estimated professional desirability: **96–99%**.

## 6. Prove the Hybrid Architecture With Real Tools

**Status: Approved migration order**

1. Invisible Unicode Cleaner
2. Duplicate Paragraph Remover
3. Punctuation Normalizer
4. Shared candidate reporting
5. Capitalization Fixer
6. MetaDataSuite

For each migration:

- Preserve its current Word-facing behavior.
- Keep VBA responsible for Preview, navigation, revalidation, Apply, and Undo.
- Compare VBA and C# candidate results.
- Preserve fast Word `Find`/`ReplaceAll` paths where they are still superior.
- Do not remove the VBA implementation until parity and rollback are proven.

Estimated combined difficulty: **55–72%**.
Estimated professional desirability: **95–99%**.

## 7. Complete the Remaining Beta Product Features

**Status: Required by the Beta contract**

- Add context-aware straight-to-curly quotes and apostrophes.
- Evaluate other useful opposite directions as modes within existing tools.
- Keep ambiguous inverse transformations review-only.
- Complete the nonhighlighted-findings review interface.
- Close out MetaDataSuite and Final Review documentation and live validation.
- Reconfirm the delivered Capitalization requirements against its checklist.
- Reconcile the launcher and tool language with the final Beta architecture.

Estimated combined difficulty: **52–68%**.
Estimated professional desirability: **94–99%**.

## 8. Build the State-Aware Hybrid Installer

**Status: Required by the Beta contract**

The same Setup executable must offer:

- **Install** when absent;
- **Update**, **Repair**, and **Uninstall** when an older version exists;
- **Repair/Reinstall** and **Uninstall** when the current version exists.

It must also:

- install a matched `.dotm`, C# engine, rules, and protocol version;
- verify hashes and compatibility;
- back up before replacement;
- retain the newest three backups;
- preserve intentional user settings when appropriate;
- remove temporary payloads during repair and uninstall;
- leave `Normal.dotm`, Word security, and trusted locations untouched;
- disable Apply and recommend Repair when engine/template versions conflict.

Estimated combined difficulty: **58–72%**.
Estimated professional desirability: **97–99%**.

## 9. Decide the Remaining Proposed Professional Corrections

**Status: Recommendations, not automatically approved**

- Suite-wide exact preservation of pre-existing highlights.
- Full shared candidate manifest across every tool.
- Confidence tiers for language-sensitive tools.
- Reference/style-aware list normalization.
- Run-aware formatting and font cleanup.
- Expanded protected/skipped reporting.
- Tool-specific changes from the professional audit.
- Generic definition-driven tool forms.
- Later named-pipe communication.
- Broader launcher reorganization and tool consolidation.

These items must be reviewed individually or accepted as a defined package before implementation.

## 10. Complete the Beta Release Gate

**Status: Final stage**

- Change the development version to `0.9.5-beta` at the appropriate checkpoint.
- Run complete source and regression validation.
- Run contract and C# engine tests.
- Synchronize both `.docm` files and generated bundles.
- Refresh Startup and release templates.
- Run focused small-fixture and large-article Word tests.
- Run the installer state matrix in isolation.
- Verify hashes, backups, repair, uninstall, and version-mismatch behavior.
- Review documentation and update the history and handoff.
- Obtain Chris's explicit authorization before committing, tagging, pushing, or publishing.

Estimated combined difficulty: **30–45%**.
Estimated professional desirability: **99–100%**.

## Overall Estimate

Estimated combined difficulty: **72–84%**.
Estimated professional desirability: **98–100%**.

## Governing Records

- [Current handoff](../WHERE_WE_LEFT_OFF.md)
- [Project history and architectural record](CleanupSuite_Project_History_and_Architectural_Record.md)
- [0.9.5 Beta contract](0.9.5-beta-contract.md)
- [0.9.5 Jump Board](0.9.5-jump-board.md)
- [Versioning and roadmap](../VERSIONING.md)
