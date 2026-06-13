# CleanupSuite for Microsoft Word — User Manual

**Documentation release target:** 0.6.5  
**Codebase documented:** 0.6.4 code-freeze candidate  
**Commit documented:** `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`  
**Prepared:** 2026-06-12

> Documentation basis: written from the pushed GitHub `0.6.4` code-freeze candidate, commit `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`. The planned `0.6.5` release should use the same codebase plus this documentation, except for documentation-blocking bug fixes if required.


## 1. What CleanupSuite does

CleanupSuite is a Microsoft Word cleanup toolkit for documents that have accumulated formatting, spacing, punctuation, layout, metadata, or pasted-content problems. It is built for careful cleanup work: preview first, verify what will change, and then apply.

Use CleanupSuite when you need to clean text copied from PDFs, web pages, email, OCR, older templates, or multiple Word files. It can normalize invisible characters, punctuation, spacing, capitalization, lists, paragraph structure, tables, breaks, links, metadata, headers/footers, footnotes/endnotes, and embedded objects.

## 2. Installation overview

The current distributable is:

```text
VBA_Cleanup_tool.txt
```

The file is pasted into a standard VBA module and installed by running:

```vba
InstallCleanupSuite
```

The installer creates the CleanupSuite forms and helper modules inside the active macro-enabled Word document or template. In the developer workflow, source is maintained in `src/`, assembled with `python assemble.py`, and the generated `VBA_Cleanup_tool.txt` is tested in Word.

### Basic installation steps

1. Open the Word file or template where CleanupSuite should live.
2. Open the VBA editor with Alt+F11.
3. Insert a standard module.
4. Paste the full contents of `VBA_Cleanup_tool.txt` into that module.
5. Run `InstallCleanupSuite`.
6. Save the file as a macro-enabled document or template, such as `.docm` or `.dotm`.
7. Close and reopen Word if ribbon changes were installed.

### Trust settings

If installation reports that controls or forms could not be created, Word probably blocked programmatic access to the VBA project. Check:

```text
File > Options > Trust Center > Trust Center Settings > Macro Settings
```

Enable:

```text
Trust access to the VBA project object model
```

A developer install may also require the VBA Extensibility reference:

```text
Microsoft Visual Basic for Applications Extensibility 5.3
```

## 3. Opening CleanupSuite

CleanupSuite opens from the main launcher. In a ribbon-enabled setup, the Add-ins tab can contain a **Cleanup Suite** button. The launcher title is **Cleanup Suite**, and its subtitle explains: "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."

The launcher contains global settings:

- **Return to main menu after apply** — returns to the launcher after applying a previewed cleanup.
- **Return to main menu after closing individual tool menus** — reopens the launcher when a tool window is closed.
- **Auto-save before running each tool** — saves the active document before cleanup when possible.
- **Reset All** — clears preview highlighting and restores CleanupSuite defaults.

## 4. Preview-first workflow

Most tools support a preview flow. Preview mode highlights or reports the items that would change, then opens the compact **Preview Actions** bar.

The Preview Actions bar shows:

- **Preview is ON / Preview is OFF** — toggles preview state.
- **Reconfigure** — clears preview highlighting and returns to the tool menu so settings can be changed.
- **Apply** — reruns the selected tool settings without preview and commits the cleanup.
- The current tool name, such as **Capitalization Fixer**.
- A short summary line when available.

You can drag the Preview Actions bar to a convenient location. In 0.6.4, the bar remembers its position while toggling preview off and on during the same session.

### Recommended preview workflow

1. Open a tool from the Cleanup Suite main menu.
2. Select the options you want.
3. Leave preview enabled when available.
4. Run the tool.
5. Review yellow highlighting and/or the summary.
6. Choose **Reconfigure** to change settings, or **Apply** to commit the changes.
7. Use Ctrl+Z immediately if the result is not what you expected.

## 5. Whole document vs. selection

Many tools can operate on the whole document or only the current selection. To clean only part of a document, select the text first and choose the selection scope in the tool when available.

Some tools are naturally document-level, such as metadata cleanup, style cleanup, or document trim. Those tools may ignore selection scope because the operation only makes sense for the whole document.

## 6. Safety features

CleanupSuite includes several safety features:

- Preview mode before applying changes.
- A guard for unsaved documents, with a prompt to save first.
- Optional auto-save before running tools.
- A custom document property marker for interrupted cleanup operations.
- A blank-document guard before applying previewed cleanup.
- A reset option that clears preview highlighting and restores default settings.
- Immediate Word Undo support for many cleanup actions.

Even with these safeguards, work on a copy of any important document until you trust the tool and settings.

## 7. Tool guide

### Invisible Unicode Cleaner

Finds and removes invisible/special characters such as non-breaking spaces, zero-width spaces, byte order marks, soft hyphens, and non-breaking hyphens.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Punctuation Normalizer

Converts smart punctuation to plain-text equivalents, including curly quotes, dashes, ellipses, and selected custom characters.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Spacing Fixer

Corrects repeated spaces, leading/trailing paragraph spaces, spacing around punctuation, and excessive blank lines.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Capitalization Fixer

Applies smart sentence capitalization, sentence case, title case, uppercase, lowercase, or custom combinations.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### List Normalizer

Standardizes bullets, numbering, list indentation, and hyphen/asterisk pseudo-lists.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Paragraph Structure Fixer

Removes empty paragraphs, collapses extra breaks, resets direct paragraph spacing, and fixes direct indentation.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Duplicate Paragraph Detector

Finds exact, normalized, or fuzzy duplicate paragraphs and can highlight or remove duplicates while keeping the first occurrence.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Font Normalizer

Resets selected direct font overrides such as face, size, bold, italic, and color back to the paragraph style.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Table Cleaner

Removes empty rows/columns, normalizes padding/borders, strips direct table formatting, or converts tables to text.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Break Normalizer

Collapses repeated page/section breaks and can convert section break types.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Document Trim

Removes trailing empty paragraphs at the end of the document while leaving Word's required final paragraph.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Direct Formatting Stripper

Removes direct character and paragraph formatting while preserving the document's styles.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Hyperlink Remover

Removes hyperlinks while keeping visible text, optionally clearing hyperlink formatting.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Soft Return Converter

Converts Shift+Enter soft returns to paragraph marks or, for selected cases, paragraphs to soft returns.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Metadata Scrubber

Removes document properties, personal information, comments, and tracked-change history before sharing.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Style Cleanup

Removes unused custom styles and remaps common style variants back to Normal.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Footnote / Endnote Remover

Removes footnotes and/or endnotes, optionally keeping note text inline in brackets.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Header / Footer Standardizer

Standardizes or clears header/footer content across sections.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

### Remove Objects & Elements

Removes pictures, text boxes, frames, horizontal lines, HTML/ActiveX controls, hidden text, and optionally tables.

**Best practice:** Run preview first when available. Review the highlighted areas or summary before applying. Use selection scope for targeted cleanup when the tool supports it.

## 8. Suggested cleanup order

For a messy imported or assembled document, a conservative order is:

1. Save a copy of the document.
2. Run Invisible Unicode Cleaner.
3. Run Punctuation Normalizer if plain punctuation is desired.
4. Run Spacing Fixer.
5. Run Paragraph Structure Fixer.
6. Run List Normalizer if lists are messy.
7. Run Break Normalizer if page/section breaks are inconsistent.
8. Run table, font, formatting, hyperlink, or style tools as needed.
9. Run Metadata Scrubber only when preparing to share externally.
10. Run Document Trim near the end.
11. Perform a final manual review.

Stop after each major step and review the document. Do not run destructive tools, such as object removal or metadata cleanup, until you have a saved copy.

## 9. Before sharing a document externally

Before sending a document outside your organization or project:

- Save a private original first.
- Accept or reject tracked changes intentionally.
- Remove comments if they should not be shared.
- Run Metadata Scrubber.
- Check headers, footers, footnotes, hidden text, and embedded objects.
- Export or save the final version only after reviewing the cleaned document.

## 10. Troubleshooting summary

If a tool does not behave as expected:

- Close and reopen Word if a form or ribbon state looks stale.
- Confirm the document is not protected.
- Try the tool on a small selection to isolate the issue.
- Run preview before applying.
- Use Ctrl+Z immediately after an unwanted apply.
- Record the tool, settings, Word version, Windows scaling, and a screenshot for bug reports.

## 11. Developer/source notes

The project is source-first. Edit `src/forms/`, `src/modules/`, or `src/installer/installer.bas`, then run `python assemble.py` so `VBA_Cleanup_tool.txt` is rebuilt from source. Do not edit the generated distributable by hand.

The 0.6.5 documentation release should use the 0.6.4 codebase and add polished documentation. It should not add new features unless a documentation-blocking bug is found.
