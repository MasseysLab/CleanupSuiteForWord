# CleanupSuite For Word 0.9.1 Alpha

CleanupSuite For Word 0.9.1 Alpha is a focused hotfix for the first public
Alpha. It preserves the 0.9.0 tools, interface, installer behavior, and safety
model.

## Fixed

- MetaDataSuite now opens and completes its quick checks when Word represents a
  synchronized personal OneDrive document as a `d.docs.live.net` URL.
- The resolver maps the Word URL back to an existing synchronized local file
  before reading the DOCX package.
- Word's caret representation of a dot immediately before the file extension
  is repaired only when the corresponding local file actually exists.
- Safe Edit Current, Group Edit Current, and open-document detection use the
  same resolved physical path, preventing duplicate opens.

The fix was reproduced and verified with the affected OneDrive document,
`Arguement for and against God..docx`, through the installed global `.dotm`.

## Install

Close Word and run `release/CleanupSuiteForWord-Setup.exe`. Existing
CleanupSuite templates are backed up automatically, with the newest three
backups retained.

## Verification

- Source assembly: 29 builders clean.
- Regression suite: 204 passed, 3 skipped, 1,095 subtests passed.
- Installed `.dotm` quick-check audit passed on the affected synchronized
  OneDrive document.
- Release template and installer-embedded template SHA-256:
  `093a2005950f5dc3d796268e23e77eeea82958dabb98898fd7231bf0bfde627d`.

See the [CleanupSuite User Manual](../documents/CleanupSuite_User_Manual.pdf)
for normal installation and use.

Publisher: MasseysLab. Software and documentation: MIT License.
