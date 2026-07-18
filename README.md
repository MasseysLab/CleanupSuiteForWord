<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 MasseysLab -->

# CleanupSuite For Word

CleanupSuite For Word is a set of preview-first tools for cleaning pasted text, paragraphs, lists, formatting, document objects, links, review content, and editable metadata in Microsoft Word; it is for speeding up careful document cleanup, not for replacing human review or repairing every document automatically.

Close Word, download and run [CleanupSuiteForWord-Setup.exe](https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe), then reopen Word and select **CleanupSuite** on the ribbon.

## Start using it

1. Open the document you want to clean.
2. Select **CleanupSuite** on Word's ribbon.
3. Choose a tool, keep **Preview ON**, and review the highlighted changes.
4. Select **Apply** only when the preview is correct.

Work on a copy of important documents. CleanupSuite's recommended auto-save setting is on by default, but it does not replace your own backups.

The [CleanupSuite User Manual](documents/CleanupSuite_User_Manual.pdf) explains every tool and the preview workflow.

## Installation choices

### Preferred: per-user installer

The installer places `CleanupSuite.dotm` in your own Word Startup folder. It does not require administrator rights, edit `Normal.dotm`, change Word's macro-security policy, or install for other Windows accounts. Existing installations are backed up first; only the newest three installer backups are retained.

The installer and app check the stable release manifest for a newer version. CleanupSuite asks before opening an update and its launcher lets you disable periodic checks; when disabled, it explains how to turn them back on. Because the Alpha installer is not yet code-signed, Windows may show a publisher/reputation warning even though the publisher recorded by the installer is MasseysLab.

### Alternative: manual installation

Close Word, download `CleanupSuite.dotm` from the newest [GitHub release](https://github.com/MasseysLab/CleanupSuiteForWord/releases), copy it to `%APPDATA%\Microsoft\Word\STARTUP`, and reopen Word. If Windows blocked the downloaded file, open its **Properties**, select **Unblock**, and copy it again.

To uninstall manually, close Word and remove `%APPDATA%\Microsoft\Word\STARTUP\CleanupSuite.dotm`. The installer also registers a normal per-user uninstall entry in Windows Settings.

## Release

- Current release: **0.9.2 Alpha**
- Git tag: `v0.9.2-alpha`
- Publisher: **MasseysLab**

See the [0.9.2 Alpha release notes](docs/Release_Notes_v0.9.2-alpha.md) for
highlights, known Alpha limits, and the verified template hash.

Software and documentation are licensed under the [MIT License](LICENSE).

## Contributor quick start

The project is source-first. Edit readable files under `src/`; do not hand-edit the generated `VBA_Cleanup_tool.txt`.

Generated tool forms use the shared two-column guided layout documented in `CONTRIBUTING.md`; single leftover choices are centered.

```powershell
python assemble.py
python -m pytest -q
powershell -ExecutionPolicy Bypass -File .\scripts\sync_docm_code_only.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\refresh_startup_dotm.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\run_word_alpha_smoke.ps1
```

Run the Word alpha smoke runner before release and complete its remaining visual review in Word.

Build the per-user release installer only after the Startup template is refreshed:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-installer.ps1
```

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
