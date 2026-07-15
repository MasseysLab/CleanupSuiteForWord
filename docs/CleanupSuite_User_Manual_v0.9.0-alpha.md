# CleanupSuite User Manual

**Target release:** 0.9.0 Alpha (`v0.9.0-alpha`)
**Prepared:** 2026-07-14
**Audience:** normal Word users, not programmers

Copyright (c) 2026 MasseysLab. Software and documentation are licensed under the MIT License.


## What CleanupSuite is

CleanupSuite is a set of tools for cleaning up messy Microsoft Word documents. It is for documents that have been copied from websites, PDFs, emails, OCR, old templates, or many different files.

It helps fix things like weird spacing, invisible characters, messy lists, bad paragraph breaks, old links, document metadata, and unwanted objects.

The safest way to use it is:

1. Work on a copy.
2. Preview what the tool will do.
3. Review the highlights or summary.
4. Apply only when the preview looks right.
5. Use Ctrl+Z right away if the result is wrong.

## What CleanupSuite is not

CleanupSuite is not magic. It cannot understand every document perfectly. It is a helper that does repetitive cleanup work faster than a person can, but you are still the final reviewer.

Do not run destructive cleanup tools on the only copy of an important document.

## Before you use it

Make a copy of the document first. This is the most important safety rule.

Good file names:

```text
MyDocument_before_cleanup.docx
MyDocument_cleanup_test.docx
MyDocument_final_after_review.docx
```

Keep Auto-save turned on unless you have a specific reason to turn it off.

## Install CleanupSuite

### Preferred installer

1. Close every Microsoft Word window.
2. Download and run [CleanupSuiteForWord-Setup.exe](https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe).
3. Reopen Word and select **CleanupSuite** on the ribbon.

The installer works only for your Windows account and does not need administrator rights. It does not edit `Normal.dotm`, weaken macro security, or add a new trusted folder. Before replacing an older copy it makes a backup and retains the newest three.

This Alpha installer is not yet code-signed, so Windows may show a publisher or reputation warning. The release page publishes the template's SHA-256 hash for verification.

### Manual alternative

Close Word, copy `CleanupSuite.dotm` from the newest GitHub release to `%APPDATA%\Microsoft\Word\STARTUP`, and reopen Word. If Windows blocked the download, right-click it, select **Properties**, select **Unblock**, and copy it again.

## Opening CleanupSuite

Open a normal Word document and select **CleanupSuite** on the ribbon. The launcher is the main menu.

The main menu groups tools by what you are trying to fix. You do not have to know the internal code names.

Common settings:

- **Auto-save before each tool**: saves before cleanup so recovery is easier. Leave this selected for normal use.
- **Return to main menu after apply**: reopens the launcher after an applied cleanup.
- **Return to main menu after closing individual tool menus**: reopens the launcher when a tool is closed with its X.
- **Check periodically for updates**: checks at most once every seven days. Clear it to stop automatic checks; CleanupSuite immediately explains how to turn them back on.
- **Check for Updates**: performs a manual update check even when periodic checks are off.
- **Reset All**: clears preview state and restores recommended global and tool defaults.

## The Preview Actions window

After you preview a tool, CleanupSuite shows a small Preview Actions window and switches Word to Reading View for review.

The buttons usually mean:

- **Preview is ON**: preview highlighting is active. **ON** is always bold. The document cannot be edited in this state.
- **Preview is OFF**: preview highlighting has been cleared or temporarily turned off.
- **Reconfigure**: go back and change the tool settings.
- **Apply**: run the same tool settings for real.
- **Previous page / Next page**: move through the preview from this window when Word's own Reading View arrows do not respond.

You can move the Preview Actions window out of the way. Turning preview off returns to an editable view; use **Reconfigure** before creating a new preview.

Preview and Apply show brief, non-intrusive progress in Word's status bar. CleanupSuite clears that status even when a tool stops with an error.

## Whole document or selection

Some tools can clean the whole document or only the selected text.

Use selection cleanup when:

- only one section is messy
- the document contains code, tables, or special formatting elsewhere
- you want to test a tool on a small area first

Use whole-document cleanup when:

- the problem is everywhere
- you are working on a copy
- the preview looks correct

## A simple cleanup path for messy pasted documents

Try this order first:

1. Invisible Unicode Cleaner
2. Punctuation Normalizer
3. Spacing Fixer
4. Paragraph Structure Fixer
5. List Normalizer
6. Document Trim

Then use other tools only if the document needs them.

## Tool-by-tool guide

### Invisible Unicode Cleaner

Use this when copied text looks normal but search, sorting, line breaks, export, or comparison behaves oddly.

It looks for hidden characters such as non-breaking spaces, zero-width spaces, soft hyphens, byte-order marks, and non-breaking hyphens.

Use preview to check the result before applying. Preview first because the characters are hard to see by eye. A high count usually means the document came from web, PDF, OCR, or another editor.

Practical advice: Run early in a cleanup pass. It makes later spacing, punctuation, and comparison steps more predictable.

### Punctuation Normalizer

Use this when you need plain-text-friendly punctuation or consistent punctuation before export or comparison.

It looks for smart quotes, curly apostrophes, em/en dashes, ellipses, and other punctuation variants.

Use preview to check the result before applying. Preview the exact punctuation types selected. Avoid converting all punctuation when a document intentionally uses typographic punctuation for publication.

Practical advice: Use Quotes only, Dashes only, or Ellipses for a narrow fix. Use All only when the target system prefers plain ASCII-like punctuation.

### Spacing Fixer

Use this when a document has been assembled from pasted text, emails, PDFs, or multiple authors.

It looks for double spaces, paragraph-edge spaces, spaces before punctuation, inconsistent spaces after punctuation, and extra blank lines.

Use preview to check the result before applying. Preview should show whether the tool is finding true spacing problems or intentionally spaced text such as code samples or aligned tables.

Practical advice: This is usually a safe early tool. For technical documents, use selection scope around prose and avoid code blocks or fixed-width examples.

### Capitalization Fixer

Use this when text has inconsistent paragraph capitalization or headings need a consistent style.

Recommended repair uses abbreviation, acronym, name, brand, heading-style, domain, and sentence-boundary protections together. Dot-attached website/domain segments are kept lowercase (`JW.org`, `Example.com`, and `Anything.blah`), while chained initials such as `P.O.Box` remain exempt. Parenthetical wording in headings is left exactly as written by default. In Custom repair, selecting **Also title-case words inside parentheses** treats each balanced parenthetical group as its own title segment.

When **Repair likely headings in title case** is cleared, suspected headings—including ALL-CAPS headings—are left completely unchanged. This prevents partial results such as changing only the longest words in an ALL-CAPS heading. Ordinary non-heading capitalization repair continues.

Custom repair also lets you choose individual protections and add document-specific exceptions for names, brands, and house acronyms.

In **Edit custom exceptions**, enter each protected term exactly as it should appear and separate terms with semicolons, for example: `MetaDataSuite; CleanUpSuite; JW`. No programming-style names, equals signs, or value pairs are used.

Use preview to check the result before applying. Preview carefully because capitalization can affect names, acronyms, product names, legal terms, and deliberately styled headings. For Custom mode, add exceptions before previewing so the preview reflects the protection list you intend to use.

Practical advice: Use on a selection first. For long documents, fix obvious problem sections rather than running broad modes on the whole document.

### List Normalizer

Use this when lists came from email, PDFs, web pages, or hand-typed outlines.

It looks for manual bullets, typed numbering, inconsistent list indentation, and hyphen/asterisk pseudo-lists.

Use preview to check the result before applying. Preview helps confirm whether normal paragraphs are being mistaken for list items.

Practical advice: Convert hyphen/star lists before normalizing bullets. Check legal/technical numbering manually before applying numbering changes.

### Paragraph Structure Fixer

Use this when a document has strange vertical gaps, inconsistent paragraph spacing, or pasted layout clutter.

It looks for empty paragraphs, repeated paragraph breaks, direct paragraph spacing overrides, and direct indentation.

Use preview to check the result before applying. Preview can show empty paragraphs or spacing issues; review around headings, signature blocks, and intentionally spaced forms.

Practical advice: Use after spacing cleanup but before final formatting polish. Avoid whole-document fixes in forms or layout-heavy documents without preview review.

### Duplicate Paragraph Detector

Use this when a document may contain duplicated blocks or repeated clauses.

It looks for exact, normalized, or fuzzy repeated paragraphs.

Use preview to check the result before applying. Always preview first. Fuzzy matching can flag similar but intentionally different paragraphs.

Practical advice: Start with highlight-only/exact matching. Move to fuzzy only when you are searching for near-duplicates in prose, not tables or legal clauses.

### Font Normalizer

Use this when text carries inconsistent direct formatting instead of using styles.

It looks for direct font face, size, bold, italic, and color overrides.

Use preview to check the result before applying. Preview paragraphs with direct overrides. Be careful with bold, italic, and color because they may be intentional emphasis.

Practical advice: Start with font face and size only. Add bold/italic/color reset only when the document should be style-controlled.

### Table Cleaner

Use this when tables are messy after pasting or importing, or need a uniform basic appearance.

It looks for empty rows/columns, padding, direct table formatting, borders, and table-to-text conversion.

Use preview to check the result before applying. Preview especially before deleting rows/columns or converting tables to text. Tables often contain intentional blank cells.

Practical advice: Clean a copy for complex tables. Use Convert table to text only when the table structure is no longer needed.

### Break Normalizer

Use this when a document has unexpected blank pages, section jumps, or inconsistent breaks from merged files.

It looks for repeated page breaks, repeated section breaks, and section-break type conversions.

Use preview to check the result before applying. Preview around section boundaries. Break changes can affect headers, footers, page numbering, margins, and columns.

Practical advice: Use after content cleanup and before final layout review. For documents with complex headers/footers, test on a copy.

### Document Trim

Use this when the end of the document has blank space, extra pages, or leftover empty paragraphs.

It looks for extra empty paragraphs at the end of a document.

Use preview to check the result before applying. Preview should show the trailing paragraphs to remove. This tool is document-end focused, not a general blank-line cleaner.

Practical advice: Run near the end of cleanup. It is a good final housekeeping step before saving or exporting.

### Formatting Stripper / Strip Formatting

Use this when a document should return to style-based formatting without changing the underlying text.

It looks for manual character and paragraph formatting layered over Word styles.

Use preview to check the result before applying. Preview paragraphs carrying direct formatting. Removing formatting can change the visible look of emphasis or layout.

Practical advice: Use Quick clean first. Use Thorough when preserving individual bold/italic words matters. Strip everything only for a deliberate full reset.

### Hyperlink Remover

Use this when a document should keep its visible text but remove clickable links.

It looks for live hyperlinks and optional hyperlink formatting.

Use preview to check the result before applying. Preview each hyperlink before applying, especially in citations, footnotes, and reference lists.

Practical advice: Clear hyperlink formatting when you want plain text. Leave formatting if the document should still visually show where links used to be.

### Soft Return Converter

Use this when text imported from PDFs or web pages uses line breaks where real paragraphs are needed.

It looks for Shift+Enter line breaks and paragraph marks.

Use preview to check the result before applying. Preview carefully because converting paragraphs to soft returns can merge structure.

Practical advice: Most cleanup uses Soft returns to paragraphs. Use Paragraphs to soft returns only on small, intentional selections.

### MetaDataSuite

Use this before a document is shared outside its trusted editing group.

It looks for document properties, application properties, custom properties, package metadata, embedded items, and other sharing-sensitive metadata signals.

Use Refresh Quick Checks to inspect the current document first. Use Clear Sharing Properties when you want to clear populated document info fields before sharing. Final Review handles comments and tracked changes separately.

Practical advice: Save a private original first. **Safe Edit Current covers every metadata field CleanupSuite classifies as safely editable**, not only Word's five Summary Info fields. The top group contains Title, Subject, Author, Keywords, Comments, and the related Category field. Attribution and classification contains Company, Manager, Last Author, Content Status, Content Type, and Language. Custom document properties lets you add, update, or delete unlinked user properties. Linked properties, CleanupSuite settings, timestamps, revision counters, security/signatures, package structure, and generated statistics stay read-only. Detail panels use inspection language because they report findings; only explicit editor and cleanup actions save or clear data.

### Final Review

Use this when comments and tracked changes are ready to be finalized before sharing or publishing.

It can remove comments and accept tracked changes. These actions change review history and may expose the final accepted text, so preview and review the document carefully.

Practical advice: Keep a private reviewed original. Use Final Review separately from MetaDataSuite so metadata cleanup and review finalization remain deliberate choices.

### Style Cleanup

Use this when the style list is cluttered after combining templates or pasted content.

It looks for unused custom styles and common style variants.

Use preview to check the result before applying. Preview the count and names of styles affected. A style used only in headers, footers, or text boxes may need manual review.

Practical advice: Use after content cleanup, before final styling. Do not run blindly on documents where custom styles carry contractual or template meaning.

### Footnote / Endnote Remover

Use this when you need a clean text extract or a version without notes.

It looks for footnotes and/or endnotes, optionally preserving note text inline.

Use preview to check the result before applying. Preview reference marks. Removing notes changes scholarly/legal text and can affect citations.

Practical advice: Save a copy. Choose keep note text inline when the note content still matters.

### Header / Footer Standardizer

Use this when headers or footers are inconsistent after merging files or templates.

It looks for header/footer formatting or content across sections.

Use preview to check the result before applying. Review every section afterward. Headers and footers can have primary, first-page, and even-page versions.

Practical advice: Use unlink sections only when you truly want each section controlled independently. Be careful with page numbers and logos.

### Object Remover / Objects

Use this when you need to reduce clutter imported from web pages or PDFs to cleaner text.

It looks for pictures, text boxes, frames, horizontal lines, HTML/ActiveX controls, hidden text, and optionally tables.

Use preview to check the result before applying. Preview the count before deleting. Object removal can be destructive and not always fully undoable.

Practical advice: Run on a copy. Leave table deletion off unless you intentionally want to remove table content, not just table formatting.

## When something looks wrong

Stop and do not keep clicking. Try these in order:

1. Press Ctrl+Z once.
2. Close without saving if the document is badly changed.
3. Reopen your backup copy.
4. Write down which tool and settings caused the problem.
5. Test the same tool on a smaller selection.

## What to tell someone helping you

Give them:

- CleanupSuite version
- Word version
- Windows version
- which tool you used
- whether Preview was on
- whether you used whole document or selection
- what happened
- a screenshot if the issue is visual

## Updates

CleanupSuite checks at most once every seven days when periodic checks are enabled. A check reads only the small release manifest; it does not upload your document or document metadata. If an update exists, CleanupSuite asks before opening the installer link.

To stop periodic checks, clear **Check periodically for updates** in the launcher. To restore them later, select it again. **Check for Updates** always remains available for a manual check.
