# CleanupSuite Human-Friendly User Manual

**Target release:** 0.8.5
**Prepared:** 2026-06-24
**Audience:** normal Word users, not programmers

> Documentation basis: CleanupSuite 0.8.5 Alpha as currently prepared in this repository. Screenshots are intentionally deferred.


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

## Opening CleanupSuite

Open the macro-enabled CleanupSuite document or template, then open the Cleanup Suite launcher. The launcher is the main menu.

The main menu groups tools by what you are trying to fix. You do not have to know the internal code names.

Common settings:

- **Show completion review after apply**: after a cleanup is applied, CleanupSuite shows the compact completion / preview action bar.
- **Return to main menu after completion review**: after the completion review is finished, the main menu opens again.
- **Return to main menu after closing individual tool menus**: if you close a tool window, the main menu comes back.
- **Auto-save before running each tool**: saves before cleanup so recovery is easier.
- **Reset All**: restores CleanupSuite settings and clears preview state.

## The Preview Actions bar

After you run a tool in preview mode, CleanupSuite may show a small Preview Actions bar.

The buttons usually mean:

- **Preview is ON**: preview highlighting is currently active.
- **Preview is OFF**: preview highlighting has been cleared or temporarily turned off.
- **Reconfigure**: go back and change the tool settings.
- **Apply**: run the same tool settings for real.

You can drag the Preview Actions bar out of the way while reviewing the document.

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

This tool is for copied text looks normal but search, sorting, line breaks, export, or comparison behaves oddly.

It looks for hidden characters such as non-breaking spaces, zero-width spaces, soft hyphens, byte-order marks, and non-breaking hyphens.

Use preview to check the result before applying. Preview first because the characters are hard to see by eye. A high count usually means the document came from web, PDF, OCR, or another editor.

Practical advice: Run early in a cleanup pass. It makes later spacing, punctuation, and comparison steps more predictable.

### Punctuation Normalizer

This tool is for you need plain-text-friendly punctuation or consistent punctuation before export/comparison.

It looks for smart quotes, curly apostrophes, em/en dashes, ellipses, and other punctuation variants.

Use preview to check the result before applying. Preview the exact punctuation types selected. Avoid converting all punctuation when a document intentionally uses typographic punctuation for publication.

Practical advice: Use Quotes only, Dashes only, or Ellipses for a narrow fix. Use All only when the target system prefers plain ASCII-like punctuation.

### Spacing Fixer

This tool is for a document has been assembled from pasted text, emails, PDFs, or multiple authors.

It looks for double spaces, paragraph-edge spaces, spaces before punctuation, inconsistent spaces after punctuation, and extra blank lines.

Use preview to check the result before applying. Preview should show whether the tool is finding true spacing problems or intentionally spaced text such as code samples or aligned tables.

Practical advice: This is usually a safe early tool. For technical documents, use selection scope around prose and avoid code blocks or fixed-width examples.

### Capitalization Fixer

This tool is for text has inconsistent paragraph capitalization or headings need a consistent style.

It looks for sentence case, title case, uppercase, lowercase, and smart paragraph capitalization.

Use preview to check the result before applying. Preview carefully because capitalization can affect names, acronyms, product names, legal terms, and deliberately styled headings.

Practical advice: Use on a selection first. For long documents, fix obvious problem sections rather than running broad modes on the whole document.

### List Normalizer

This tool is for lists came from email, PDF, web pages, or hand-typed outlines.

It looks for manual bullets, typed numbering, inconsistent list indentation, and hyphen/asterisk pseudo-lists.

Use preview to check the result before applying. Preview helps confirm whether normal paragraphs are being mistaken for list items.

Practical advice: Convert hyphen/star lists before normalizing bullets. Check legal/technical numbering manually before applying numbering changes.

### Paragraph Structure Fixer

This tool is for a document has strange vertical gaps, inconsistent paragraph spacing, or pasted layout clutter.

It looks for empty paragraphs, repeated paragraph breaks, direct paragraph spacing overrides, and direct indentation.

Use preview to check the result before applying. Preview can show empty paragraphs or spacing issues; review around headings, signature blocks, and intentionally spaced forms.

Practical advice: Use after spacing cleanup but before final formatting polish. Avoid whole-document fixes in forms or layout-heavy documents without preview review.

### Duplicate Paragraph Detector

This tool is for a document may contain copied/pasted duplicate blocks or repeated clauses.

It looks for exact, normalized, or fuzzy repeated paragraphs.

Use preview to check the result before applying. Always preview first. Fuzzy matching can flag similar but intentionally different paragraphs.

Practical advice: Start with highlight-only/exact matching. Move to fuzzy only when you are searching for near-duplicates in prose, not tables or legal clauses.

### Font Normalizer

This tool is for text carries inconsistent direct formatting instead of using styles.

It looks for direct font face, size, bold, italic, and color overrides.

Use preview to check the result before applying. Preview paragraphs with direct overrides. Be careful with bold, italic, and color because they may be intentional emphasis.

Practical advice: Start with font face and size only. Add bold/italic/color reset only when the document should be style-controlled.

### Table Cleaner

This tool is for tables are messy after paste/import or need a uniform basic appearance.

It looks for empty rows/columns, padding, direct table formatting, borders, and table-to-text conversion.

Use preview to check the result before applying. Preview especially before deleting rows/columns or converting tables to text. Tables often contain intentional blank cells.

Practical advice: Clean a copy for complex tables. Use Convert table to text only when the table structure is no longer needed.

### Break Normalizer

This tool is for a document has unexpected blank pages, section jumps, or inconsistent breaks from merged files.

It looks for repeated page breaks, repeated section breaks, and section-break type conversions.

Use preview to check the result before applying. Preview around section boundaries. Break changes can affect headers, footers, page numbering, margins, and columns.

Practical advice: Use after content cleanup and before final layout review. For documents with complex headers/footers, test on a copy.

### Document Trim

This tool is for the end of the document has blank space, extra pages, or leftover empty paragraphs.

It looks for extra empty paragraphs at the end of a document.

Use preview to check the result before applying. Preview should show the trailing paragraphs to remove. This tool is document-end focused, not a general blank-line cleaner.

Practical advice: Run near the end of cleanup. It is a good final housekeeping step before saving or exporting.

### Formatting Stripper / Strip Formatting

This tool is for a document should return to style-based formatting without changing the underlying text.

It looks for manual character and paragraph formatting layered over Word styles.

Use preview to check the result before applying. Preview paragraphs carrying direct formatting. Removing formatting can change the visible look of emphasis or layout.

Practical advice: Use Quick clean first. Use Thorough when preserving individual bold/italic words matters. Strip everything only for a deliberate full reset.

### Hyperlink Remover

This tool is for a document should keep visible text but remove clickable links.

It looks for live hyperlinks and optional hyperlink formatting.

Use preview to check the result before applying. Preview each hyperlink before applying, especially in citations, footnotes, and reference lists.

Practical advice: Clear hyperlink formatting when you want plain text. Leave formatting if the document should still visually show where links used to be.

### Soft Return Converter

This tool is for text imported from PDFs/web pages uses line breaks where real paragraphs are needed.

It looks for Shift+Enter line breaks and paragraph marks.

Use preview to check the result before applying. Preview carefully because converting paragraphs to soft returns can merge structure.

Practical advice: Most cleanup uses Soft returns to paragraphs. Use Paragraphs to soft returns only on small, intentional selections.

### Metadata Scrubber

This tool is for a document is being shared outside the trusted editing group.

It looks for document properties, author information, comments, and tracked changes.

Use preview to check the result before applying. Preview/report what will be removed. Some privacy actions may not be meaningfully reversible after saving.

Practical advice: Save a private original first. Run before external sharing, not during active collaborative editing.

### Style Cleanup

This tool is for the style list is cluttered after combining templates or pasted content.

It looks for unused custom styles and common style variants.

Use preview to check the result before applying. Preview the count and names of styles affected. A style used only in headers, footers, or text boxes may need manual review.

Practical advice: Use after content cleanup, before final styling. Do not run blindly on documents where custom styles carry contractual or template meaning.

### Footnote / Endnote Remover

This tool is for you need a clean text extract or a version without notes.

It looks for footnotes and/or endnotes, optionally preserving note text inline.

Use preview to check the result before applying. Preview reference marks. Removing notes changes scholarly/legal text and can affect citations.

Practical advice: Save a copy. Choose keep note text inline when the note content still matters.

### Header / Footer Standardizer

This tool is for headers/footers are inconsistent after merging files or templates.

It looks for header/footer formatting or content across sections.

Use preview to check the result before applying. Review every section afterward. Headers and footers can have primary, first-page, and even-page versions.

Practical advice: Use unlink sections only when you truly want each section controlled independently. Be careful with page numbers and logos.

### Object Remover / Objects

This tool is for you need to reduce pasted/web/PDF clutter to cleaner text.

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

## Screenshot placeholders for later

Images to add later:

- Main menu
- A simple tool window
- Preview highlighting in the document
- Preview Actions bar
- Success or summary message
- Trust Center setting if needed
