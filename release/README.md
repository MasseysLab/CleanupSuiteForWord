# Release Artifacts

`CleanupSuiteForWord-Setup.exe` is the preferred per-user installer. Its stable GitHub link is:

<https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe>

The stable installer embeds and verifies the frozen 0.9.3 standalone VBA-only
`CleanupSuite.dotm`, then installs it in the current user's Word Startup folder.
It never changes `Normal.dotm`, Word macro-security policy, or machine-wide
settings. Before replacement it saves a backup and retains the newest three.
The same Setup offers Install, Update, Repair/Reinstall, and Uninstall according
to the detected installation.

Build the stable release files only from the verified frozen template:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-vba-only-installer.ps1 `
  -TemplatePath .\build\v0.9.3-vba-default\CleanupSuite.dotm
```

Publisher: MasseysLab. Software and documentation: MIT License.

See `docs/Release_Notes_v0.9.3-beta-vba-final.md` for stable release highlights,
verification details, and the retained internal version identifier.

## Publishing order

1. Merge the verified release commit into `main` and confirm the Build and Test
   workflow passes.
2. Preserve the frozen `v0.9.3-beta-vba-final` tag and its verified
   `CleanupSuite.dotm`; the technical tag is historical identity, not a statement
   that 0.9.3 is an unstable channel.
3. Attach `CleanupSuiteForWord-Setup.exe` to that GitHub Release with its exact
   filename and publish both Setup and template SHA-256 values.
4. Keep the versioned user-manual PDF on the release page so it is
   self-contained.
5. Verify the stable Setup link, `latest-release.ini`, template download, and a
   manual Check for Updates from Word after publication.
