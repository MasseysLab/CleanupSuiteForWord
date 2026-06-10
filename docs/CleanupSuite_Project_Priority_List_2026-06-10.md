# CleanupSuite Project Priority List

Date: 2026-06-10

This is the working priority list for turning CleanupSuite from a working Word VBA toolkit into a polished, global Word add-in/template.

The current direction is clear:

- keep the project useful and stable first
- redesign the launcher before touching every individual form
- build the preview/action model before deep tool-form polish
- keep aiming toward a global `.dotm` template/add-in once behavior is solid

## Current Reality

CleanupSuite is already a broad suite of 19 Word cleanup tools. The biggest remaining gap is not tool count; it is product flow.

The current system has:

- a working source pipeline: `src` -> `assemble.py` -> `VBA_Cleanup_tool.txt`
- generated/installed VBA forms controlled heavily by `src/installer/installer.bas`
- a working `.docm` project file for testing
- a practice folder with test documents
- design notes for the approved modern `C`-vibe form direction
- known caution around rerunning the full installer and around running tools against the macro host document itself

Important project rule:

Do not treat `VBA_Cleanup_tool.txt` as the normal editing surface. Edit structured source under `src`, then run `python assemble.py`.

## P0 - Protect The Work

These are the safety tasks to keep the project recoverable before heavier UI changes.

- Push current committed design/spec work to GitHub when convenient.
- Keep the existing backup zip and practice folder intact.
- Before major VBA edits, make a fresh backup zip of the project folder.
- Do not clean up or revert unrelated working-tree changes unless they are deliberately reviewed.
- Continue testing on normal target documents, not on `CleanupSuite_Workspace.docm` itself, whenever possible.
- Keep the crash/reset mitigation in `modCleanupHelpers.bas`: do not auto-save or stamp custom document properties when the active document is the macro host.

Why this matters:

The project has enough moving parts now that the risk is not one bad code line; it is losing track of which copy is canonical.

## P1 - Launcher Redesign

Do this first.

Goal:

Make the suite feel organized and intentional the moment it opens.

Recommended launcher categories:

- Text and Characters
- Paragraphs, Breaks, and Lists
- Layout and Document Structure
- Formatting, Links, and Styles
- Review, Privacy, and Removals

Launcher design priorities:

- Replace the flat two-column button grid with category-based tool rows or compact cards.
- Put each tool name beside a one-line description.
- Integrate help into each tool row/card instead of separate tiny help buttons.
- Keep the launcher dense enough that it still feels like a working utility, not a landing page.
- Show destructive or privacy-related tools with clearer caution language.
- Keep the top-level tool list stable for now; do not split tools yet.

Acceptance checks:

- A user can scan the launcher and understand where to start.
- `Document Cleanup` is not used as a vague category.
- Table, break, trim, metadata, footnote, and object tools no longer feel misplaced.
- The launcher still opens all 19 existing tools correctly.
- `python assemble.py` passes after source changes.

## P2 - Action Model After Preview

Do this immediately after the launcher foundation.

Goal:

Stop making the user inspect document highlights while the full suite/form blocks the document.

Recommended flow:

1. User configures a tool in the main form.
2. User clicks `Preview`.
3. The main form closes, hides, or minimizes out of the way.
4. A small modeless control panel remains available.
5. User reviews the document and chooses the next action.

Compact preview panel actions:

- `Apply`
- `Close Preview`
- `Undo` when reliable
- `Back to Options` when useful
- `Stop` only for tools that actually need interruption support

Use this especially for:

- Duplicate Paragraph Detector
- Formatting Stripper
- Object Remover
- Unicode Cleanup
- Paragraph Structure Fixer
- Table Cleaner

Acceptance checks:

- Preview no longer feels like a buried checkbox.
- The document remains easy to inspect after preview.
- Destructive actions are deliberate follow-up choices.
- The pattern can be reused without rewriting every tool from scratch.
- Modeless form behavior is tested carefully in Word.

## P3 - Shared Form Language

Do this before redesigning every individual tool.

Goal:

Create one repeatable form style so each tool does not become a one-off redesign.

Shared form structure:

- header with tool name and one-sentence purpose
- decision panels grouped by intent
- scope panel
- preview/run action area
- visible help entry

Visual rules:

- Use the approved `C` vibe as the base.
- Keep the restraint of an Office-adjacent utility.
- Avoid decorative fake-web styling that VBA cannot support well.
- Use panels, spacing, labels, and button hierarchy more than visual effects.

Implementation target:

Most of this likely belongs in `src/installer/installer.bas`, because that file currently creates, styles, captions, and lays out controls.

Acceptance checks:

- Forms have consistent title, spacing, action order, and scope language.
- Long captions fit without awkward wrapping.
- Button order is consistent.
- Help is visible but no longer dominates the layout.
- Tab order and default buttons feel sensible.

## P4 - Complex Tool Form Redesign

Do these before the simpler tools, because they will prove the design system.

First complex-tool candidates:

- Invisible Unicode Cleaner
- Duplicate Paragraph Detector
- Direct Formatting Stripper
- Object Remover
- Header / Footer Standardizer
- Table Cleaner

Tool-specific notes:

- Unicode Cleanup is the best reference form for the approved `C` vibe.
- Duplicate Detector should detect first and remove later.
- Object Remover needs stronger separation between objects, hidden text, and table deletion.
- Header / Footer Standardizer needs `Standardize` and `Clear all` to feel like clearly different modes.
- Table Cleaner needs a distinct conversion area for single-column table-to-text behavior.

Acceptance checks:

- Each complex form is easier to understand than the current checkbox stack.
- Risky actions are visually separated.
- Preview and apply behavior match the new action model.
- Existing cleanup logic still works.
- `python assemble.py` passes after each small batch.

## P5 - Simple Tool Form Redesign

Do after the complex tools prove the system.

Tools:

- Document Trim
- Hyperlink Remover
- Soft Return Converter
- Font Normalizer
- Spacing Fixer
- Capitalization Fixer
- List Normalizer
- Paragraph Structure Fixer
- Punctuation Normalizer
- Style Cleanup
- Footnote / Endnote Remover
- Metadata Scrubber

Design goal:

Use a lighter version of the same form language. Do not make small tools feel oversized.

Acceptance checks:

- Simple forms remain quick.
- The user can run common defaults without thinking too much.
- Scope and preview behavior are consistent with the larger forms.

## P6 - Sub-tool Reorganization

Do not start here.

Investigate again after launcher and preview behavior are working.

Likely later changes:

- Split `Metadata Scrubber` into `Privacy Scrubber` and `Review Markup Finalizer`.
- Move `Hidden text` out of `Object Remover` and into the privacy/review family.
- Move table deletion out of generic object removal.
- Rename `Footnote / Endnote Remover` to `Footnote / Endnote Tools` if it keeps both delete and inline-convert behavior.

Near-term decision:

Keep the current top-level tool list stable while improving categories, descriptions, and risk treatment.

## P7 - Global Word Add-in / Template

Do after the suite is stable and the main workflows feel right.

Goal:

Make CleanupSuite usable on any Word document as a global `.dotm` add-in/template.

Considerations:

- Keep testing in `.docm` while behavior is still changing.
- Move toward `.dotm` when the launcher, preview model, and major forms are stable.
- Confirm startup/ribbon behavior works cleanly from Word's Startup folder or trusted add-in location.
- Make sure the suite acts on the active document, not the add-in/template file.
- Keep the current source pipeline as the canonical authoring path.

Acceptance checks:

- Opening Word with the template installed exposes CleanupSuite reliably.
- The launcher works on any active document.
- Cleanup tools do not modify or save the global template by accident.
- Install/uninstall instructions are clear enough for a normal user.

## P8 - Testing And Release Discipline

Keep this running alongside every implementation stage.

Core checks:

- Run `python assemble.py` after source changes.
- Test each changed tool in Preview mode first.
- Test against the practice folder documents.
- Run a Word compile after meaningful VBA batches.
- Verify no unexpected UserForm components are left behind.
- Keep a note of any manual Word behavior that cannot be validated through scripts.

Manual smoke-test order:

1. Open a normal target document.
2. Launch CleanupSuite.
3. Confirm launcher layout and categories.
4. Open every tool once.
5. Run Preview for each changed tool.
6. Apply on a copy only.
7. Confirm Word does not close, reset, or show VBA break mode.
8. Confirm Undo behavior is understandable.

## Things To Avoid For Now

- Do not redesign every form before the launcher and preview model are settled.
- Do not split tools just because the current labels feel imperfect.
- Do not build the `.dotm` global template before the active-document safety model is proven.
- Do not rely on the full installer rerun until the blocking dialog behavior is understood.
- Do not hand-edit the distributable as the primary workflow.

## Recommended Next Move

Start with P1: Launcher Redesign.

After that, immediately build P2: Action Model After Preview.

Those two pieces create the spine of the product. Once they feel right, the individual forms can be modernized without repeatedly changing direction.
