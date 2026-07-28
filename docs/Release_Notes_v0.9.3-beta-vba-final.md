# CleanupSuite For Word 0.9.3 Stable — Final VBA-Only Release

**Tag:** `v0.9.3-beta-vba-final`
**Internal version identifier:** `0.9.3-beta-vba-final`
**Distribution:** Stable per-user Setup and manual-install `CleanupSuite.dotm`
**Stable installer channel:** `0.9.3`

## Purpose

This stable release freezes the last complete CleanupSuite build that runs
entirely in VBA. It preserves a recoverable, distributable checkpoint before any
real tool is made dependent on the approved C#/VBA hybrid engine.

## Highlights

- Completes the structural-safety package. Ordinary cleanup tools protect Word
  table cells, row markers, final document markers, nested tables, merged or
  ambiguous structures, and revalidate destructive candidates immediately before
  Apply.
- Makes Table Cleaner the sole owner of row and column deletion. Structural
  choices default to OFF, confirmation defaults to No, Preview and Apply share the
  same eligibility rules, and the complete operation remains one Word Undo action.
- Redesigns Duplicate Paragraph Remover around Exact, Normalized, and Fuzzy
  matching, with a separately demarcated Include empty paragraphs option and
  protected/removable structural counts.
- Expands minimal Preview markers to actions that previously showed counts without
  a visible location. Invisible boundaries use the nearest safe one-character cue
  instead of broad page or paragraph shading.
- Replaces the two Preview navigation buttons with four equal controls:
  `[Change] [Page] ◀ Previous | Next ▶ [Page] [Change]`.
  Page buttons turn one page; Change buttons jump to the nearest page containing
  at least one detected change.
- Preserves the faster display-only punctuation highlighting and grouped Apply
  path on large documents.
- Keeps Reading View stable while Preview is active, including reliable backward
  page repainting without reintroducing the old punctuation flicker.
- Uses exact restoration records so Preview removes only the temporary formatting
  that CleanupSuite added.
- Shows the launcher warning: “It is dangerous to turn off auto save!”

## Validation

- Python regression suite: **214 tests passed, 3 skipped**.
- Generated VBA: **29 builders clean**, 17,978 lines, 1,234,163 bytes.
- Preview lifecycle smoke: **PASS**.
- Structural-safety smoke and fixture validation: **PASS**.
- Word Alpha smoke: all automated checks **PASS**; the expected visual-review item
  was completed manually.
- Live regression on the 53-page Argument for God document:
  - Punctuation Preview completed in approximately 35 seconds.
  - Minimal punctuation markers remained visible.
  - Change Next moved to the next page containing changes.
  - Change Previous visibly returned to the prior changed page.
  - All four equal-size colored navigation buttons remained stable.
  - The disposable test copy was closed without saving.
- The optional PDF export step in the structural smoke hung in the current Word
  automation environment. The same structural smoke and fixture passed when that
  optional export was skipped; the live document review was completed directly in
  Word. No CleanupSuite logic failure was observed.

## Installation

### Preferred: stable Setup

1. Close Word.
2. Download `CleanupSuiteForWord-Setup.exe` from this release or the stable link
   in the repository README.
3. Confirm its SHA-256 matches the published Setup value and scan it with
   Microsoft Defender.
4. Run Setup and choose Install, Update, Repair/Reinstall, or Uninstall as
   appropriate.
5. Reopen Word.

### Manual alternative

If Setup cannot be used, download `CleanupSuite.dotm` from this release, verify
the template hash below, scan it, close Word, and copy it to
`%APPDATA%\Microsoft\Word\STARTUP`. If Windows marked the file as downloaded,
open **Properties**, select **Unblock**, and then select **Apply** and **OK**.

See the [unsigned-install safety guide](Unsigned_Installation_Guide.md). Do not
disable SmartScreen, Defender, Word macro security, or the Trust Center.

The stable Setup embeds this exact frozen template and the stable update manifest
now points to it. The future hybrid engine installer remains separate work for
the official `0.9.5-beta`.

Verified `CleanupSuite.dotm` SHA-256:
`4F815EE0DF4BACBDA5150A82DB2D66739592FA2098184C8D7EC7B535D73AB10C`

Verified stable `CleanupSuiteForWord-Setup.exe` SHA-256:
`B9A98AA9EE609821F5EEA348F7B1CD3325CCE0D6B154408BBDAF9ABB0B6E0AB9`

## Deferred

- Reverse/inverse cleanup modes were explicitly deferred from this standalone
  release.
- The C# sidecar, engine handshake, real-tool hybrid integration, and state-aware
  hybrid installer begin after this checkpoint.

Estimated combined difficulty: **58–70%**.
Estimated professional desirability: **98–100%**.
