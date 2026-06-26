# Capitalization Smart Repair Follow-Ups

This document records the next versions of the deterministic capitalization repair tool beyond the current minimal first version.

## Highest-Value Next Steps

- Add a small user-editable custom exception list for names, brands, and house-style acronyms.
- Generate larger compiled data packs for abbreviations, acronyms, and protected mixed-case terms from plain-text source lists.
- Improve style-aware heading detection so built-in Word heading styles influence the repair pass directly.
- Add document-level caching so repeated previews and applies do not rebuild lookup dictionaries every run.

## Later Refinements

- Expand abbreviation coverage for legal, medical, academic, and corporate writing.
- Add optional regional packs and specialty vocabularies that can be compiled into separate bundles.
- Improve sentence-boundary handling around nested quotes, numbered outlines, and OCR-damaged punctuation.
- Add telemetry-style debug output for difficult paragraphs so heuristic tuning is easier in future releases.
