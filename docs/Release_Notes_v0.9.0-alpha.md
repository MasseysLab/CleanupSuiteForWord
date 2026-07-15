# CleanupSuite For Word 0.9.0 Alpha

CleanupSuite For Word 0.9.0 Alpha is the first packaged public Alpha of the
preview-first Word cleanup suite. It combines the global Word template,
per-user installer, update checks, test documents, and complete user manual in
one release.

## Highlights

- Twenty focused tools for pasted text, paragraphs, lists, formatting, document
  objects, links, review content, and editable metadata.
- Reading View previews with a modal Preview Actions panel, panel-only page
  navigation, and a clear separation between Preview and Apply.
- Faster Capitalization repair with protected acronyms, names, domains,
  possessives, quotations, headings, and friendly custom exceptions.
- MetaDataSuite auditing plus grouped safe editing of all metadata fields
  classified as safely editable; audit-only findings remain read-only.
- Shared, nonintrusive progress reporting with guaranteed cleanup.
- Per-user installer that backs up an existing template and retains the newest
  three backups without changing `Normal.dotm` or machine-wide settings.
- Periodic update checks that can be disabled, with a manual check always
  available from the launcher.

## Install

Close Word, run `release/CleanupSuiteForWord-Setup.exe`, reopen Word, and select
**CleanupSuite** on the ribbon. The installer does not require administrator
rights. Manual Startup-folder installation remains documented in the
[README](../README.md).

## Alpha notes

- The installer is not yet code-signed, so Windows may show a publisher or
  reputation warning.
- CleanupSuite supports desktop Microsoft Word for Windows and does not alter
  Word's macro-security policy.
- Style Cleanup and Break Normalizer remain clearly separated as risky Alpha
  tools and should be previewed especially carefully.
- CleanupSuite accelerates careful cleanup; it does not replace document
  backups or human review.

## Verification

- Source assembly: 29 builders clean.
- Regression suite: 200 passed, 3 skipped.
- Word Alpha and Preview Lifecycle smoke checks passed.
- Release template, manifest, and installer-embedded template SHA-256:
  `8c25d6189300cc48c8c16f7214e52741f235cc339cb32b005367424023738b5e`.

See the [CleanupSuite User Manual](../documents/CleanupSuite_User_Manual.pdf)
for installation, preview, tool, reset, and troubleshooting instructions.

Publisher: MasseysLab. Software and documentation: MIT License.
