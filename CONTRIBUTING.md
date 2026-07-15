<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 MasseysLab — see LICENSE for full terms. -->

# Contributing to CleanupSuite For Word

CleanupSuite is source-first. Edit the readable files under `src`; never edit
the generated `VBA_Cleanup_tool.txt` bundle by hand. The detailed tool-authoring
reference is the
[CleanupSuite Programmer's Guide For Adding Tools](docs/CleanupSuite_Programmers_Guide_For_Adding_Tools_v0.9.0-alpha.md).

## Development quick start

```powershell
python assemble.py
python -m unittest discover -s tests
powershell -ExecutionPolicy Bypass -File .\scripts\sync_docm_code_only.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\run_word_alpha_smoke.ps1
```

The assembler validates the builder structure and refreshes both generated VBA
bundles. Word COM smoke tests are still required because source checks cannot
prove that an installed VBA project compiles and behaves correctly.

## Project map

- `src/forms/` — readable UserForm code
- `src/modules/` — launcher, update, metadata, preview, progress, and safety services
- `src/installer/` — the VBA self-installer and generated-control definitions
- `src/manifest.txt` — generated-bundle assembly order
- `scripts/sync_docm_code_only.ps1` — synchronizes source into both `.docm` files
- `scripts/refresh_startup_dotm.ps1` — refreshes the installed global template
- `installer/` — per-user Windows installer source and build script
- `release/` — publishable installer, template, and stable update manifest
- `tests/` — source, packaging, and release regression tests

## Adding a tool

1. Add the readable form module under `src/forms/`.
2. Add its `BUILDER` entry to `src/manifest.txt` in the intended assembly order.
3. Add the form, controls, captions, and mirrored layout to
   `src/installer/installer.bas`.
4. Add the launcher button, description, risk label, and main-menu Help action.
   Normal tool forms do not contain their own Help or Apply buttons.
5. Add the form to the developer uninstall component list.
6. Update `scripts/sync_docm_code_only.ps1` when an existing embedded form needs
   newly created controls before refreshed code can position them.
7. Add regression tests and user documentation appropriate to the new behavior.
8. Assemble, synchronize, compile, and exercise the tool in Word on a disposable
   document before considering it complete.

## Shared form contract

Generic tools use the shared **two-column guided layout**. In particular:

- normal opt/chk choices are paired into two columns
- Custom choices are still two-column choices
- single leftover choices are centered
- the body follows an intro-first guided layout
- runtime forms call `LayoutCleanupToolForm Me`
- installed forms mirror the result through `LayoutGenericToolControls`

Reset is a small utility in the upper-right title row. Preview is the visible
tool action. Apply appears only in the modal Preview Actions window after the
user has reviewed the document. Scope choices sit beside Preview at the bottom
when the tool supports selected text.

Read [Tool-Form Rules](docs/Tool-Form%20Rules.md) before changing shared layout
behavior. A form must remain within the `768 x 510` VBA-point outer envelope and
must fit a 1024 x 768 pixel display after accounting for the native title bar.

## Safety and coding rules

- Reject protected documents before making changes.
- Use `GuardBeforeCleanup`, `MarkCleanupStart`, and `MarkCleanupEnd` consistently.
- Use the shared progress service; do not write directly to `Application.StatusBar`.
- Keep Preview analysis inside the Preview path so Apply does not repeat it.
- Preserve distinct contained-range, overlapping-range, and outside-table scope
  semantics by using the shared paragraph helpers.
- Preserve paragraph and table-cell markers with `ParagraphBodyText` and
  `ReplaceParagraphBodyText`.
- Group an Apply operation in one Word `UndoRecord` and guarantee cleanup on
  normal and error exits.
- Iterate backward when deleting collection members or stored document ranges.
- Never identify deletions by highlight color; users may already have matching
  highlights in their documents.
- Use explicit option-button `GroupName` values because generated controls are
  flat on the form.
- Prefer descriptive names and one action per line. Add shared helpers only when
  current code uses them immediately, and remove replaced duplicates in the same
  change.
- Avoid unsupported Word members such as
  `ClearCharacterDirectFormatting` and `ClearParagraphDirectFormatting`; the
  suite uses `Font.Reset` and `ParagraphFormat.Reset`.

## Required verification

Before submitting a change:

1. Run `python assemble.py`.
2. Run `python -m pytest -q` or `python -m unittest discover -s tests`.
3. Run `check.ps1`.
4. Synchronize both `.docm` files when VBA source changed.
5. Run the Word Alpha and Preview Lifecycle smoke checks when forms, preview,
   progress, or shared helpers changed.
6. Confirm `git diff --check` is clean and generated bundles match source.
7. Leave Word closed before rebuilding or replacing the Startup template.

Do not commit local build output, rendered QA pages, smoke reports, backup files,
or restart packets. Publish release artifacts only from `release/` and the stable
user-manual PDFs under `documents/`.

By contributing, you agree that your contribution is licensed under the
[MIT License](LICENSE).
