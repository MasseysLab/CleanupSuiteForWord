# CleanupSuite Release Notes — 0.8.0 to 0.8.5

**Release:** 0.8.5 Alpha
**Prepared:** 2026-06-24
**Branch:** codex-0.8.0-release

## Summary

0.8.5 is the first Alpha release. It refines the guided-form layout introduced in 0.8.0 with a two-column info box, cleaner option captions, and a suppressed preview form title.

---

## Changes

### Two-column guided info box

The info box displayed when a user hovers or focuses on a tool option was redesigned from a single merged line to a two-column layout.

- The left column shows a concise display name for the control (e.g. "Smart quotes", "Extra blank lines").
- The right column shows a plain-language advisory explaining what the option does or what to watch out for.
- Row heights are now dynamic — rows with longer advisory text grow to fit.
- Column width adjusts based on the length of the display name.
- Display names and advisory text are looked up through two new helper functions: `GuidedInfoDisplayName()` and `GuidedInfoAdvisory()`, which cover all current tool controls.
- Controls without a specific advisory entry fall back to a template-generated description from `GuidedGenericAdvisory()`.

This change is in `modCleanupHelpers.bas`. The installer was updated to match.

### Option caption simplification

The "Custom (choose below)" caption on the Custom option button in all applicable guided forms was shortened to "Custom". The contextual note that options appear below is no longer needed given the current layout.

Affected forms: `frmListCleanup`, `frmParagraphCleanup`, `frmPunctuationCleanup`, `frmSpacingCleanup`, `frmUnicodeCleanup`.

### Preview form title suppressed

The Preview Actions bar title bar is now explicitly blanked (`Me.Caption = ""`). This prevents the default UserForm name from appearing in the title bar of the preview panel.

### Option height adjustments

Default option control heights were bumped slightly (16 → 18, 24 → 28, 38 → 44 pts) to give long-label options more breathing room.

### Long-caption width handling

`GuidedSingleChoiceWidth()` now widens option pairs automatically when a caption exceeds 30 or 42 characters, preventing text truncation on longer option labels.

---

## Not changed

- No tool behavior changes.
- No tool set additions or removals.
- No documentation set restructure (that is scoped to 0.9.0).

---

## Verification

```text
python assemble.py
python vbaeval.py validate
python -m unittest discover -s tests
git diff --check
```

Plus Word smoke test per `docs/Alpha_Readiness_Checklist_0.8.5.md`.
