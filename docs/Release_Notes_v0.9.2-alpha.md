# CleanupSuite For Word 0.9.2 Alpha

CleanupSuite For Word 0.9.2 Alpha is the final focused Alpha hotfix before the
0.9.5 Beta development cycle. It preserves the existing tool set, installer,
and safety model while correcting preview performance and presentation.

## Fixed

- Punctuation Normalizer preview now uses Word's display-only highlight engine
  instead of applying document formatting one match at a time.
- Punctuation Apply now processes selected punctuation families with grouped
  wildcard `ReplaceAll` passes, avoiding slow per-match replacement loops on
  long documents.
- Preview remains in Reading View with visible yellow punctuation highlights;
  closing, applying, or reconfiguring restores the original view and highlight
  visibility.
- The punctuation preview no longer toggles screen updating, eliminating its
  extra full-document repaints and flicker.
- Paragraph Fixer now marks only the affected paragraph boundary when previewing
  paragraph spacing, rather than shading the entire text area.
- Setup bypasses cached release metadata so the stable installer immediately
  retrieves the newly published hotfix template.

The punctuation fixes were reproduced against the large synchronized OneDrive
document `Arguement for and against God..docx`. The all-types preview completed
with 276 matches. A final installed-template check confirmed yellow highlights
and a stable Reading View.

## Install

Close Word and run `release/CleanupSuiteForWord-Setup.exe`. Existing
CleanupSuite templates are backed up automatically, with the newest three
backups retained.

## Verification

- Source assembly: 29 builders clean.
- Regression suite: 200 passed, 3 skipped.
- Live punctuation preview: stable Reading View with yellow highlights.
- Isolated installer install/uninstall: exit codes 0; installed template hash
  matched the release template and uninstall removed it.
- Release template and installer-embedded template SHA-256:
  `a256a27bb58583c163687c45dc70dcb9074212d73a85d027b33521be48cf9cc0`.

See the [CleanupSuite User Manual](../documents/CleanupSuite_User_Manual.pdf)
for normal installation and use.

Publisher: MasseysLab. Software and documentation: MIT License.
