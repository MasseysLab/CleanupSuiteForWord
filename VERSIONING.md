# CleanupSuite Versioning

CleanupSuite is pre-1.0 while the suite is still being shaped into a polished global Word add-in/template.

The project uses a simple SemVer-style ladder, but with `0.x.0` milestones until the add-in is ready for normal release.

## Current Version

`0.6.0`

This means the working suite baseline now includes the category-based launcher redesign.

## Roadmap

| Version | Meaning |
| --- | --- |
| `0.5.0` | Working suite baseline, safety fixes, design direction, project tracking, and versioning discipline. |
| `0.6.0` | Launcher redesign with better categories, tool descriptions, and help placement. |
| `0.7.0` | Post-preview action model with compact review controls such as Apply, Close Preview, Undo, and Back to Options. |
| `0.8.0` | Complex tool form redesigns using the approved modern panel style. |
| `0.9.0` | Simple tool form polish, consistency pass, help/caption cleanup, and release-candidate testing. |
| `1.0.0` | Global `.dotm` Word add-in/template release suitable for normal use on any document. |

## Version Rules

- Update `SUITE_VERSION` in `src/modules/modCleanupLauncher.bas` when a milestone is reached.
- Rebuild `VBA_Cleanup_tool.txt` with `python assemble.py` before distributing a version.
- Mention the version in backup zip names when the backup represents a milestone.
- Keep Git commits small enough that each milestone can be understood later.
- Do not call the project `1.0.0` until the global add-in/template workflow is stable.

## Naming Examples

- `CleanupSuite-v0.5.0-working-copy-2026-06-10.zip`
- `CleanupSuite-v0.6.0-launcher-redesign.zip`
- `CleanupSuite-v1.0.0-global-addin.zip`
