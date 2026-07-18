# CleanupSuite Versioning

CleanupSuite is pre-1.0 while the suite is still being shaped into a polished global Word add-in/template.

The project uses a simple SemVer-style ladder, but with `0.x.0` milestones until the add-in is ready for normal release.

## Current Version

`0.9.2-alpha`

This Alpha hotfix preserves the packaged tool set while correcting preview performance and presentation: punctuation previews use a fast display-only highlight path, bulk punctuation Apply avoids per-match replacement loops, Reading View remains stable, and paragraph-spacing previews mark only the affected boundary. Git tag: `v0.9.2-alpha`.

## Roadmap

| Version | Meaning |
| --- | --- |
| `0.5.0` | Working suite baseline, safety fixes, design direction, project tracking, and versioning discipline. |
| `0.6.0` | Launcher redesign with better categories, tool descriptions, and help placement. |
| `0.6.4` | Code freeze / final programming-side release candidate. |
| `0.6.5` | Documentation-complete release using the same codebase as `0.6.4`, except for documentation-blocking bug fixes if required. |
| `0.7.0` | Next programming milestone begins. |
| `0.7.4` | Code freeze / final programming-side release candidate for the `0.7.x` line. |
| `0.7.4.5` | Small bugfix bridge for release-blocking Preview Actions and compile issues before `0.7.5`. |
| `0.7.5` | Documentation-complete release using the same codebase as `0.7.4`, except for documentation-blocking bug fixes if required. |
| `0.8.0` | Complex tool form redesigns using the approved modern panel style. |
| `0.8.5` | Internal pre-release milestone used while the Alpha was being completed. |
| `0.9.0-alpha` | Official Alpha with the global template, installer, update checks, and full documentation. Git tag: `v0.9.0-alpha`. |
| `0.9.1-alpha` | Alpha hotfix for MetaDataSuite on synchronized personal OneDrive documents. Git tag: `v0.9.1-alpha`. |
| `0.9.2-alpha` | Alpha hotfix for punctuation performance and Preview/Reading View presentation. Git tag: `v0.9.2-alpha`. |
| `0.9.5-beta` | Official Beta release after Alpha feedback, full MetaData-project incorporation, and the planned Capitalization follow-up upgrades. |
| `1.0.0` | Global `.dotm` Word add-in/template release suitable for normal use on any document. |

## Version Rules

- Update `SUITE_VERSION` in `src/modules/modCleanupLauncher.bas` when a milestone is reached.
- Rebuild `VBA_Cleanup_tool.txt` with `python assemble.py` before distributing a version.
- Mention the version in backup zip names when the backup represents a milestone.
- Keep Git commits small enough that each milestone can be understood later.
- After `0.6.4` is pushed, no feature or code changes go into `0.6.5` unless required to correct documentation-blocking bugs.
- After `0.7.4` is pushed, no feature or code changes go into `0.7.5` unless required to correct documentation-blocking bugs.
- Treat the `0.9.x-alpha` line as Alpha and reserve the entire `0.9.5` milestone for Beta.
- Do not remove historical documentation or release history without an explicit maintenance decision.
- Do not call the project `1.0.0` until the global add-in/template workflow is stable.

## Naming Examples

- `CleanupSuite-v0.5.0-working-copy-2026-06-10.zip`
- `CleanupSuite-v0.6.0-launcher-redesign.zip`
- `CleanupSuite-v1.0.0-global-addin.zip`
