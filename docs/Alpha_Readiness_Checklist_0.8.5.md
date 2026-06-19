# CleanupSuite 0.8.5 Alpha Readiness Checklist

**Target:** 0.8.5 Alpha
**Created:** 2026-06-19

## Purpose

This checklist defines the next real milestone after the preserved `0.8.0` release state.

The goal for `0.8.5` is Alpha readiness: a working release that Chris can use and evaluate with confidence, without doing broad house-cleaning or full documentation cleanup early.

## Alpha Scope

- Confirm the current tool-form layout works well enough across representative simple, custom, and special-case tools.
- Confirm the main launcher, User Manual button, Help buttons, Preview, Apply, Reset, and return-to-main flows behave consistently.
- Confirm generated source, practice bundle, and both `.docm` files stay synced.
- Confirm `SUITE_VERSION` and release-facing docs identify the project as `0.8.5` when the Alpha bump happens.
- Preserve historical docs and old branches until the `0.9.0` documentation milestone.

## Alpha Verification

Run before calling `0.8.5` ready:

```text
python assemble.py
python vbaeval.py validate
python -m unittest discover -s tests
git diff --check
```

Also perform a Word smoke test on the real `.docm` files:

- open the main launcher
- open Document Trim as the simplest guided form
- open Invisible Unicode Cleaner as a Custom-mode form
- open Capitalization Fixer as the closest working reference form
- open Header / Footer Standardizer as the special-case layout form
- run at least one Preview flow
- confirm Apply and Reset are reachable and visually clean
- confirm no duplicate tool title appears
- confirm there is no large wasted vertical space
- confirm Custom divider lines appear and disappear correctly

## Not For 0.8.5

- Do not clean house.
- Do not prune historical `0.6.x` or `0.7.x` docs.
- Do not delete old release history just to make the tree look tidy.
- Do not rewrite the full documentation set unless it blocks Alpha use.

Those belong with `0.9.0`, which is Alpha plus full documentation.

## Likely 0.8.5 Closeout

- Bump `SUITE_VERSION` from `0.8.0` to `0.8.5`.
- Rebuild generated bundles.
- Sync the practice bundle.
- Sync both `.docm` files.
- Update current-facing docs and handoff references from `0.8.0` to `0.8.5` where appropriate.
- Run final verification.
- Commit and push on the current release branch or a dedicated `0.8.5` branch, depending on Chris's preference at that time.
