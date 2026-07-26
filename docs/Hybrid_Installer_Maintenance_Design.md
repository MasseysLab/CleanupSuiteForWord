# CleanupSuite Hybrid Installer and Maintenance Design

**Target:** `0.9.5-beta`  
**Status:** implemented in source and proven through the isolated maintenance matrix; official signing and release publication remain gated.

## One Setup, One Matched Package

CleanupSuite uses one per-user Setup executable for installation and continuing
maintenance. The Setup embeds one matched, offline package:

- `CleanupSuite.dotm`;
- the self-contained `CleanupSuite.Engine.exe`;
- Contract 1.0 protocol definitions;
- Contract 1.0 operation rules; and
- `installation-manifest.json`, which binds every component by version, relative
  path, byte length, and SHA-256.

The engine carries its own runtime. A user does not install, select, update, or
maintain .NET separately.

The Setup package is the update unit. Word's update check obtains a newer Setup;
Setup itself does not fetch a mixture of components during installation. This
prevents a transient network or release-pointer change from creating a
template/engine mismatch halfway through maintenance.

## Visible State Matrix

| Detected state | Visible maintenance choices | Default |
| --- | --- | --- |
| Absent | Install | Install |
| Older or legacy | Update, Repair/Reinstall, Uninstall | Update |
| Current and healthy | Repair/Reinstall, Uninstall | Repair/Reinstall |
| Current but incomplete or hash-mismatched | Repair/Reinstall, Uninstall | Repair/Reinstall |
| Newer than this Setup | Uninstall; use the newer Setup for repair | Close |

An older Setup does not silently downgrade a newer installation.

`Repair/Reinstall` means: verify this Setup's embedded package, back up the
existing template, replace the matched components, write the manifest last, and
verify the complete installed result. On an older installation, Repair/Reinstall
uses the package represented by the Setup being run and therefore may bring the
installation to that Setup's version; the visible Update action communicates that
version transition more directly.

## Per-User Paths

- Word template:
  `%APPDATA%\Microsoft\Word\STARTUP\CleanupSuite.dotm`
- Product files:
  `%LOCALAPPDATA%\MasseysLab\CleanupSuiteForWord`
- Engine:
  `Engine\CleanupSuite.Engine.exe`
- Contract material:
  `Contracts\Hybrid\v1`
- Matched manifest:
  `installation-manifest.json`
- Installer backups:
  `Backups`

The installer does not require administrator rights and does not modify
`Normal.dotm`, Word macro-security policy, the Trust Center, trusted locations, or
machine-wide registry settings.

## Safety and Recovery

- Word must be closed for real installation or maintenance.
- Official hybrid builds require an Authenticode signing certificate.
- The release build signs the engine before hashing and embedding it, then signs
  the completed Setup.
- Every embedded and staged component is checked against the package manifest.
- Existing matched files are captured for rollback before the first replacement.
- A failure restores the exact pre-maintenance files and removes the staging
  directory.
- The existing template is backed up before replacement; only the newest three
  installer backups are retained.
- The manifest is written after the engine, rules, and template, so an incomplete
  write cannot advertise a completed match.
- A complete post-install hash check is required before success is reported.
- CleanupSuite user settings are preserved through Install, Update, Repair, and
  Uninstall.
- Uninstall removes the template, engine, contract files, manifest, and Windows
  uninstall registration while retaining the newest three backups and user
  settings.

The installed Setup cache may remain briefly when Windows is executing that same
file during uninstall. It is not a runtime component and does not leave
CleanupSuite loaded in Word.

## Word Runtime Trust Boundary

The VBA bridge retains an explicit environment-variable path only for repository
development tests. Without that override, it reads the installed matched
manifest, requires the exact suite and protocol versions, recognizes only the
four approved component identities and paths, verifies byte lengths and SHA-256
hashes, and completes the engine capability handshake before analysis.

Any missing, modified, mismatched, or unknown component fails closed. No
candidate is accepted and no Apply begins. The user is directed to run Setup and
choose **Repair/Reinstall**.

## Isolated Evidence

`scripts\run_hybrid_installer_matrix.ps1` runs the real built Setup against a
disposable `build\installer-matrix-*` root. It covers:

- absent inspection and Install;
- legacy and older-version detection plus Update;
- current healthy detection;
- intentional engine damage and Repair recommendation;
- intentional failure after engine replacement, exact rollback, and staging
  cleanup;
- successful Repair/Reinstall;
- repeated repair with newest-three backup retention;
- Uninstall;
- component removal; and
- preservation of a user-settings sentinel and installer backups.

The test root suppresses registry writes, signature enforcement, and the
real-Word-process gate. It never touches the user's Startup folder or installed
CleanupSuite files.

A second disposable Word smoke synchronizes a real `.dotm`, packages and installs
it into an isolated root, removes the development engine override, and exercises
the production manifest resolver. It passed manifest/version/path/length/hash
validation, the engine capability handshake, seven-candidate Unicode analysis,
and stale-scope rejection. A live visual review also confirmed clean absent,
older, current, and damaged Setup layouts at the machine's active display scale.

Estimated combined difficulty: **72–84%**.  
Estimated professional desirability: **98–99.7%**.
