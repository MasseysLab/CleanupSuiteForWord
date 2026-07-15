# CleanupSuite Tool-Form Rules

Use these rules before changing a CleanupSuite tool form. Shared layout changes
need careful reasoning because a small control-placement change can affect every
installed form and the supported 1024 x 768 pixel display size.

**Current 0.9.0 Alpha Course**

1. Treat `0.9.0-alpha` as the current release and reserve `0.9.5-beta` for the
   next public milestone.
2. Current release documentation consists of:
   `CleanupSuite_User_Manual_v0.9.0-alpha.md`,
   `CleanupSuite_Programmers_Guide_For_Adding_Tools_v0.9.0-alpha.md`, and the
   stable PDF under `documents/CleanupSuite_User_Manual.pdf`.
3. Current release export artifacts belong in `documents/` and must be listed
   in `docs/README_DOC_PACKAGE.md`.
4. Historical documentation remains a historical record. Do not rename or
   rewrite it merely to make the current release list look smaller.
5. Keep release-facing rules aligned with `README.md`, `VERSIONING.md`, and the
   current programmer guide.

**Preview-Only Apply Rule**

1. Normal guided cleanup tool-forms configure settings and run Preview only.
2. Apply appears only in the Preview Actions window after a preview has run.
3. Existing internal apply procedures may remain as hidden plumbing for `RunAfterPreview`, but normal tool-forms must not expose a visible Apply button.
4. Reset is a small secondary utility in the top-right title row, not a bottom workflow button.
5. Scoped tools use one compact bottom row: `Selected text only` above `Entire document` on the left, and a large Preview button centered vertically on the right.
6. Full-document-only tools show one prominent Preview button at the bottom.
7. MetaDataSuite is an advanced dashboard exception and does not follow the generic guided cleanup layout.

**Tool Form Layout**

0. Hyperlink Cleaner is the current example tool design to reference for guided
   tool-forms. Match its compact two-mode choice layout, per-choice risk chips,
   selected-choice emphasis in the information box, and preview-only action flow
   unless a tool has a clear reason to differ.
1. All tool-forms follow the same sectional order: top-left body title with the
   small Reset utility at the right, centered bold instruction, main choices,
   optional divider and custom choices, optional Select All / Deselect All row,
   informational box, then the Preview and Scope area.
2. The form title bar should not carry the tool title. The visible tool title
   belongs in the top-left body area.
3. The centered bold sentence is the instruction or intro, not a repeated title.
4. Main choices use two columns. A lone item on a row is centered only when it is
   truly a single row item, such as a fifth option or a one-choice tool.
5. Controls align by their checkbox or option-button glyphs. True center-row
   controls are centered on the form, not merely nudged inward.
6. Custom is the last main choice when a tool has a Custom mode.
7. When Custom is selected, a horizontal divider appears below the main choices,
   followed by the custom choices in the same two-column style.
8. Select All and Deselect All appear only when a visible multi-select custom
   section is active. They do not appear for plain single-choice rows.
9. The informational box lists the currently relevant choice group. Main-choice
   mode lists main choices; Custom mode lists custom choices. Selected items are
   bolded.
10. Scoped tools place scope choices on the left with `Selected text only` above
    `Entire document`, and Preview on the right centered vertically in that row.
11. Tools that always work on the full document do not show an Entire document
    option and use one prominent Preview button at the bottom.
12. Apply is not visible inside normal guided tool-forms. Apply is only visible
    in the Preview Actions window after preview. Help stays on the main menu,
    not inside the tool-forms.
13. Reset is a tiny secondary utility in the top-right title row. It is not a
    bottom workflow button.
14. Header / Footer Standardizer special rule: Unlink sections appears before Set
    alignment and shares its row. When Set alignment is selected, Left and Right
    share one row, and Center appears on the next row truly centered.
15. Keep vertical spacing compact. Do not leave large gaps between the choices,
    information box, scope, or Preview unless content needs the room.
16. When generated bundles or root project files change, update the matching
    practice files in `Practice - Try CleanupSuite Here` before finishing.
