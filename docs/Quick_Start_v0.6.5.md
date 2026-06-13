# CleanupSuite Quick Start

**Documentation release target:** 0.6.5  
**Codebase documented:** 0.6.4 code-freeze candidate  
**Commit documented:** `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`

> Documentation basis: written from the pushed GitHub `0.6.4` code-freeze candidate, commit `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`. The planned `0.6.5` release should use the same codebase plus this documentation, except for documentation-blocking bug fixes if required.


## Install

1. Open the Word document or template where CleanupSuite will live.
2. Open the VBA editor with Alt+F11.
3. Insert a standard module.
4. Paste the contents of `VBA_Cleanup_tool.txt`.
5. Run `InstallCleanupSuite`.
6. Save as `.docm` or `.dotm`.

If Word blocks installation, enable **Trust access to the VBA project object model** in Trust Center settings.

## First safe cleanup

1. Open a copy of a document.
2. Open Cleanup Suite.
3. Pick a tool, such as Spacing Fixer.
4. Leave preview enabled.
5. Run the tool.
6. Review the highlights or summary.
7. Click **Reconfigure** to change settings or **Apply** to commit.
8. Use Ctrl+Z immediately if the result is wrong.

## Good first tool order

1. Invisible Unicode Cleaner
2. Punctuation Normalizer
3. Spacing Fixer
4. Paragraph Structure Fixer
5. List Normalizer
6. Document Trim

## Main safety rules

- Work on a copy until you trust the cleanup.
- Preview before applying.
- Use selection scope when possible.
- Do not run destructive removal tools on your only copy.
- Keep Auto-save on unless you have a reason to turn it off.
