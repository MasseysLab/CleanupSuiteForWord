# CleanupSuite Tool Reference - Refined Tool Sections

**Target release:** 0.8.5
**Prepared:** 2026-06-24
**Repo:** `MasseysLab/CleanupSuiteForWord`

> Documentation basis: CleanupSuite 0.8.5 Alpha as currently prepared in this repository. This reference is meant to reflect the current tool set and current guided-form behavior without pretending that screenshots or every export artifact are already finished.


## Purpose

This document replaces vague repeated wording such as "run preview first" with tool-specific guidance. It is meant to be folded into the user manual and help text over time.

## General cleanup rule

For most documents, think in this order:

1. Make a copy.
2. Remove hidden/invisible text problems.
3. Normalize punctuation and spacing.
4. Fix paragraph/list/table structure.
5. Clean formatting, links, metadata, and removable elements.
6. Review visually.
7. Save/export only after review.

## Current interface note

Most CleanupSuite tool forms now use a shared guided layout with:

- a body title instead of a repeated native form title
- a centered instruction line
- compact spacing
- a full-row Preview button
- Apply and Reset as the only bottom action buttons
- a two-column info box that shows a concise option name on the left and plain-language advisory text on the right when a control is focused

That layout polish changes presentation, but the tool guidance below is still about what each tool does and how to use it safely.

## Tool sections

### Invisible Unicode Cleaner

**What it fixes:** hidden characters such as non-breaking spaces, zero-width spaces, soft hyphens, byte-order marks, and non-breaking hyphens.

**Use it when:** copied text looks normal but search, sorting, line breaks, export, or comparison behaves oddly.

**Preview guidance:** Preview first because the characters are hard to see by eye. A high count usually means the document came from web, PDF, OCR, or another editor.

**Best workflow:** Run early in a cleanup pass. It makes later spacing, punctuation, and comparison steps more predictable.

### Punctuation Normalizer

**What it fixes:** smart quotes, curly apostrophes, em/en dashes, ellipses, and other punctuation variants.

**Use it when:** you need plain-text-friendly punctuation or consistent punctuation before export/comparison.

**Preview guidance:** Preview the exact punctuation types selected. Avoid converting all punctuation when a document intentionally uses typographic punctuation for publication.

**Best workflow:** Use Quotes only, Dashes only, or Ellipses for a narrow fix. Use All only when the target system prefers plain ASCII-like punctuation.

### Spacing Fixer

**What it fixes:** double spaces, paragraph-edge spaces, spaces before punctuation, inconsistent spaces after punctuation, and extra blank lines.

**Use it when:** a document has been assembled from pasted text, emails, PDFs, or multiple authors.

**Preview guidance:** Preview should show whether the tool is finding true spacing problems or intentionally spaced text such as code samples or aligned tables.

**Best workflow:** This is usually a safe early tool. For technical documents, use selection scope around prose and avoid code blocks or fixed-width examples.

### Capitalization Fixer

**What it fixes:** sentence case, title case, uppercase, lowercase, and smart paragraph capitalization.

**Use it when:** text has inconsistent paragraph capitalization or headings need a consistent style.

**Preview guidance:** Preview carefully because capitalization can affect names, acronyms, product names, legal terms, and deliberately styled headings.

**Best workflow:** Use on a selection first. For long documents, fix obvious problem sections rather than running broad modes on the whole document.

### List Normalizer

**What it fixes:** manual bullets, typed numbering, inconsistent list indentation, and hyphen/asterisk pseudo-lists.

**Use it when:** lists came from email, PDF, web pages, or hand-typed outlines.

**Preview guidance:** Preview helps confirm whether normal paragraphs are being mistaken for list items.

**Best workflow:** Convert hyphen/star lists before normalizing bullets. Check legal/technical numbering manually before applying numbering changes.

### Paragraph Structure Fixer

**What it fixes:** empty paragraphs, repeated paragraph breaks, direct paragraph spacing overrides, and direct indentation.

**Use it when:** a document has strange vertical gaps, inconsistent paragraph spacing, or pasted layout clutter.

**Preview guidance:** Preview can show empty paragraphs or spacing issues; review around headings, signature blocks, and intentionally spaced forms.

**Best workflow:** Use after spacing cleanup but before final formatting polish. Avoid whole-document fixes in forms or layout-heavy documents without preview review.

### Duplicate Paragraph Detector

**What it fixes:** exact, normalized, or fuzzy repeated paragraphs.

**Use it when:** a document may contain copied/pasted duplicate blocks or repeated clauses.

**Preview guidance:** Always preview first. Fuzzy matching can flag similar but intentionally different paragraphs.

**Best workflow:** Start with highlight-only/exact matching. Move to fuzzy only when you are searching for near-duplicates in prose, not tables or legal clauses.

### Font Normalizer

**What it fixes:** direct font face, size, bold, italic, and color overrides.

**Use it when:** text carries inconsistent direct formatting instead of using styles.

**Preview guidance:** Preview paragraphs with direct overrides. Be careful with bold, italic, and color because they may be intentional emphasis.

**Best workflow:** Start with font face and size only. Add bold/italic/color reset only when the document should be style-controlled.

### Table Cleaner

**What it fixes:** empty rows/columns, padding, direct table formatting, borders, and table-to-text conversion.

**Use it when:** tables are messy after paste/import or need a uniform basic appearance.

**Preview guidance:** Preview especially before deleting rows/columns or converting tables to text. Tables often contain intentional blank cells.

**Best workflow:** Clean a copy for complex tables. Use Convert table to text only when the table structure is no longer needed.

### Break Normalizer

**What it fixes:** repeated page breaks, repeated section breaks, and section-break type conversions.

**Use it when:** a document has unexpected blank pages, section jumps, or inconsistent breaks from merged files.

**Preview guidance:** Preview around section boundaries. Break changes can affect headers, footers, page numbering, margins, and columns.

**Best workflow:** Use after content cleanup and before final layout review. For documents with complex headers/footers, test on a copy.

### Document Trim

**What it fixes:** extra empty paragraphs at the end of a document.

**Use it when:** the end of the document has blank space, extra pages, or leftover empty paragraphs.

**Preview guidance:** Preview should show the trailing paragraphs to remove. This tool is document-end focused, not a general blank-line cleaner.

**Best workflow:** Run near the end of cleanup. It is a good final housekeeping step before saving or exporting.

### Formatting Stripper / Strip Formatting

**What it fixes:** manual character and paragraph formatting layered over Word styles.

**Use it when:** a document should return to style-based formatting without changing the underlying text.

**Preview guidance:** Preview paragraphs carrying direct formatting. Removing formatting can change the visible look of emphasis or layout.

**Best workflow:** Use Quick clean first. Use Thorough when preserving individual bold/italic words matters. Strip everything only for a deliberate full reset.

### Hyperlink Remover

**What it fixes:** live hyperlinks and optional hyperlink formatting.

**Use it when:** a document should keep visible text but remove clickable links.

**Preview guidance:** Preview each hyperlink before applying, especially in citations, footnotes, and reference lists.

**Best workflow:** Clear hyperlink formatting when you want plain text. Leave formatting if the document should still visually show where links used to be.

### Soft Return Converter

**What it fixes:** Shift+Enter line breaks and paragraph marks.

**Use it when:** text imported from PDFs/web pages uses line breaks where real paragraphs are needed.

**Preview guidance:** Preview carefully because converting paragraphs to soft returns can merge structure.

**Best workflow:** Most cleanup uses Soft returns to paragraphs. Use Paragraphs to soft returns only on small, intentional selections.

### Metadata Scrubber

**What it fixes:** document properties, author information, comments, and tracked changes.

**Use it when:** a document is being shared outside the trusted editing group.

**Preview guidance:** Preview/report what will be removed. Some privacy actions may not be meaningfully reversible after saving.

**Best workflow:** Save a private original first. Run before external sharing, not during active collaborative editing.

### Style Cleanup

**What it fixes:** unused custom styles and common style variants.

**Use it when:** the style list is cluttered after combining templates or pasted content.

**Preview guidance:** Preview the count and names of styles affected. A style used only in headers, footers, or text boxes may need manual review.

**Best workflow:** Use after content cleanup, before final styling. Do not run blindly on documents where custom styles carry contractual or template meaning.

### Footnote / Endnote Remover

**What it fixes:** footnotes and/or endnotes, optionally preserving note text inline.

**Use it when:** you need a clean text extract or a version without notes.

**Preview guidance:** Preview reference marks. Removing notes changes scholarly/legal text and can affect citations.

**Best workflow:** Save a copy. Choose keep note text inline when the note content still matters.

### Header / Footer Standardizer

**What it fixes:** header/footer formatting or content across sections.

**Use it when:** headers/footers are inconsistent after merging files or templates.

**Preview guidance:** Review every section afterward. Headers and footers can have primary, first-page, and even-page versions.

**Best workflow:** Use unlink sections only when you truly want each section controlled independently. Be careful with page numbers and logos.

### Object Remover / Objects

**What it fixes:** pictures, text boxes, frames, horizontal lines, HTML/ActiveX controls, hidden text, and optionally tables.

**Use it when:** you need to reduce pasted/web/PDF clutter to cleaner text.

**Preview guidance:** Preview the count before deleting. Object removal can be destructive and not always fully undoable.

**Best workflow:** Run on a copy. Leave table deletion off unless you intentionally want to remove table content, not just table formatting.

## Future screenshot placeholders

Later documentation should add screenshots for:

- the main launcher
- a basic text cleanup tool
- a preview highlight example
- the Preview Actions bar
- a more destructive tool warning example
- the final success/report message
