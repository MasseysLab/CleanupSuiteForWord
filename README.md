<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 MasseysLab -->

# CleanupSuite For Word

CleanupSuite For Word is a set of preview-first tools for cleaning pasted text, paragraphs, lists, formatting, document objects, links, review content, and editable metadata in Microsoft Word; it is for speeding up careful document cleanup, not for replacing human review or repairing every document automatically.

## Install the stable release

Close Word, download and run the stable
[CleanupSuiteForWord-Setup.exe](https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe),
then reopen Word and select **CleanupSuite** on the ribbon. The default Setup
installs **CleanupSuite 0.9.3 Stable**, the final standalone VBA-only release.
It detects an existing installation and offers the appropriate Install, Update,
Repair/Reinstall, or Uninstall choices.

The historical `0.9.2-alpha` Setup is no longer the default download.

## Start using it

1. Open the document you want to clean.
2. Select **CleanupSuite** on Word's ribbon.
3. Choose a tool, keep **Preview ON**, and review the highlighted changes.
4. Select **Apply** only when the preview is correct.

Work on a copy of important documents. CleanupSuite's recommended auto-save setting is on by default, but it does not replace your own backups.

The [CleanupSuite User Manual](documents/CleanupSuite_User_Manual.pdf) explains every tool and the preview workflow.

## Installation choices

### Preferred: per-user installer

The stable installer places the `0.9.3` standalone VBA-only
`CleanupSuite.dotm` in your own Word Startup folder. It does not require
administrator rights, edit `Normal.dotm`, change Word's macro-security policy,
or install for other Windows accounts. Existing installations are backed up
first; only the newest three installer backups are retained.

The installer and app check the **stable release channel** for a newer version.
CleanupSuite asks before opening an update, and its launcher lets you disable
periodic checks; when disabled, it explains how to turn them back on. Because
the current stable installer is not yet code-signed, Windows may show a
publisher/reputation warning. Follow the [unsigned-install safety
guide](docs/Unsigned_Installation_Guide.md) to verify and scan the official
download before deciding whether to run it. Never disable SmartScreen, Microsoft
Defender, Word macro security, or the Trust Center to install CleanupSuite.

### Manual-install alternative

If Setup cannot be used, close Word, download
[CleanupSuite.dotm](https://github.com/MasseysLab/CleanupSuiteForWord/releases/download/v0.9.3-beta-vba-final/CleanupSuite.dotm)
from the official stable release, verify and scan it, copy it to
`%APPDATA%\Microsoft\Word\STARTUP`, and reopen Word. If Windows blocked the
downloaded file, open its **Properties**, select **Unblock**, and copy it again.
Future hybrid releases must use their matched Setup package because the
template, analysis engine, protocol, and rules are installed together.

To uninstall manually, close Word and remove `%APPDATA%\Microsoft\Word\STARTUP\CleanupSuite.dotm`. The installer also registers a normal per-user uninstall entry in Windows Settings.

## Releases

- Stable installer release:
  [**0.9.3 Stable — Final VBA-Only Release**](https://github.com/MasseysLab/CleanupSuiteForWord/releases/tag/v0.9.3-beta-vba-final)
  (`v0.9.3-beta-vba-final`)
- Historical release: **0.9.2 Alpha** (`v0.9.2-alpha`)
- Publisher: **MasseysLab**

The stable Setup and manual template both install the same verified 0.9.3
standalone VBA-only release. See the
[0.9.3 Stable release notes](docs/Release_Notes_v0.9.3-beta-vba-final.md) for
its safety, navigation, validation, and installation details.

Software and documentation are licensed under the [MIT License](LICENSE).

## Code signing policy

CleanupSuite is pursuing sponsored public-trust signing for its open-source
releases. Free code signing provided by [SignPath.io](https://signpath.io/),
certificate by [SignPath Foundation](https://signpath.org/), is the preferred
no-cost path, subject to application approval.

See the complete [code signing policy](CODE_SIGNING_POLICY.md), [privacy
policy](PRIVACY.md), and [unsigned-install safety
guide](docs/Unsigned_Installation_Guide.md).

## Contributor quick start

The project is source-first. Edit readable files under `src/`; do not hand-edit the generated `VBA_Cleanup_tool.txt`.

Generated tool forms use the shared two-column guided layout documented in `CONTRIBUTING.md`; single leftover choices are centered.

```powershell
python -m pip install -r requirements-dev.txt
python assemble.py
python -m pytest -q
powershell -ExecutionPolicy Bypass -File .\scripts\sync_docm_code_only.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\refresh_startup_dotm.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\run_word_alpha_smoke.ps1
```

Run the Word alpha smoke runner before release and complete its remaining visual review in Word.

Build the stable 0.9.3 VBA-only Setup only from the verified frozen release
template:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-vba-only-installer.ps1 `
  -TemplatePath .\build\v0.9.3-vba-default\CleanupSuite.dotm
```

`installer\build-installer.ps1` is reserved for the future matched hybrid
package and must not replace the stable VBA-only Setup before that release gate
passes.

Important paths:

| Path | Purpose |
| --- | --- |
| `src/manifest.txt` | Source assembly order |
| `src/installer/installer.bas` | VBA component/control installer and layout mirror |
| `src/modules/` | Shared VBA services and launcher entry points |
| `src/forms/` | UserForm code |
| `scripts/sync_docm_code_only.ps1` | Sync source into macro-enabled Word artifacts |
| `installer/` | Per-user Windows Setup source and build script |
| `release/` | Setup EXE, global template, and stable release manifest |
| `tests/` | Python regression suite |

See [CONTRIBUTING.md](CONTRIBUTING.md), [VERSIONING.md](VERSIONING.md), and the [CleanupSuite Programmer's Guide For Adding Tools](docs/CleanupSuite_Programmers_Guide_For_Adding_Tools_v0.9.0-alpha.md) for the full development workflow.
