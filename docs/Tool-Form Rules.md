I raised your Reasoning level to High.

Use these rules before touching any Cleanup Suite tool-form. If the model or
reasoning level is unknown at the start of a design-heavy tool-form discussion,
ask Chris. For broad shared-form/layout passes, recommend a stronger model and
High reasoning.

**Current 0.8.5 Alpha Course**

1. Treat `0.8.5` as the current working release unless Chris explicitly changes
   the target.
2. The active working branch is `codex-0.8.0-release`. The old
   `codex-0.7.0-start` name is historical and should not be used for new work.
3. Keep `AI_HANDOFF.md` short and current. Put long session history in
   `AI_HANDOFF_ARCHIVE.md` instead of letting the active handoff become hard to
   use again.
4. Current release docs belong in the `0.8.5` set:
   `Human_Friendly_User_Manual_v0.8.5.md`,
   `Programmers_Guide_Adding_Tools_v0.8.5.md`,
   `Tool_Reference_Refined_v0.8.5.md`,
   `Release_Notes_0.8.0_to_0.8.5.md`, and
   `Documentation_Handoff_0.8.5.md`.
5. Current release export artifacts belong in `documents/` and should be listed
   in `docs/README_DOC_PACKAGE.md`.
6. Use `scripts/export_markdown_docs_to_docx.py` when refreshing DOCX exports
   from the Markdown docs.
7. Historical `0.6.x` and `0.7.x` docs should remain as historical records.
   Do not rename or rewrite them just to make a current release list look tidy.
8. When changing release docs or export artifacts, update
   `docs/README_DOC_PACKAGE.md`, `docs/Documentation_Handoff_0.8.5.md`, and
   `docs/Release_Notes_0.8.0_to_0.8.5.md` if their file lists or claims change.
9. `0.8.5` is the Alpha release. `0.9.0` is Alpha plus full documentation.
   Clean-house work waits until the `0.9.0` documentation state is complete
   unless Chris explicitly asks for earlier cleanup.

**Tool Form Layout**

1. All tool-forms follow the same sectional order: top-left body title, centered
   bold instruction, main choices, optional divider and custom choices, optional
   Select All / Deselect All row, informational box, Preview button, Scope row
   when scope is available, then Apply and Reset.
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
10. The Preview button covers the full margined row.
11. Scope appears below Preview when the tool supports selection scope. Tools that
    always work on the full document do not show an Entire document option.
12. Apply and Reset are the only bottom action buttons on tool-forms. Help stays
    on the main menu, not inside the tool-forms.
13. Header / Footer Standardizer special rule: Unlink sections appears before Set
    alignment and shares its row. When Set alignment is selected, Left and Right
    share one row, and Center appears on the next row truly centered.
14. Keep vertical spacing compact. Do not leave large gaps between the choices,
    information box, Preview, Scope, or bottom buttons unless content needs the
    room.
15. When generated bundles or root project files change, update the matching
    practice files in `Practice - Try CleanupSuite Here` before finishing.
