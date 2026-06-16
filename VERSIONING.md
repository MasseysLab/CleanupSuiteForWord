# CleanupSuite Versioning

CleanupSuite is pre-1.0 while the suite is still being shaped into a polished global Word add-in/template.

The project uses a simple SemVer-style ladder, but with `0.x.0` milestones until the add-in is ready for normal release.

## Current Version

`0.8.0`

This is the current programming milestone, carrying forward the locked `0.7.5` release with the shared Preview Actions polish and release-ready UI cleanup as the starting point for the `0.8.x` line.

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
| `0.8.5` | Official Alpha release. |
| `0.9.0` | Simple tool form polish, consistency pass, help/caption cleanup, and release-candidate testing. |
| `0.9.5` | Official Beta release. |
| `1.0.0` | Global `.dotm` Word add-in/template release suitable for normal use on any document. |

## Version Rules

- Update `SUITE_VERSION` in `src/modules/modCleanupLauncher.bas` when a milestone is reached.
- Rebuild `VBA_Cleanup_tool.txt` with `python assemble.py` before distributing a version.
- Mention the version in backup zip names when the backup represents a milestone.
- Keep Git commits small enough that each milestone can be understood later.
- After `0.6.4` is pushed, no feature or code changes go into `0.6.5` unless required to correct documentation-blocking bugs.
- After `0.7.4` is pushed, no feature or code changes go into `0.7.5` unless required to correct documentation-blocking bugs.
- Do not call the project `1.0.0` until the global add-in/template workflow is stable.

## Naming Examples

- `CleanupSuite-v0.5.0-working-copy-2026-06-10.zip`
- `CleanupSuite-v0.6.0-launcher-redesign.zip`
- `CleanupSuite-v1.0.0-global-addin.zip`
