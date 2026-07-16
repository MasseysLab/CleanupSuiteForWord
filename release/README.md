# Release Artifacts

`CleanupSuiteForWord-Setup.exe` is the preferred per-user installer. Its stable GitHub link is:

<https://github.com/MasseysLab/CleanupSuiteForWord/raw/refs/heads/main/release/CleanupSuiteForWord-Setup.exe>

The installer checks `latest-release.ini`, verifies the published SHA-256 hash, and installs `CleanupSuite.dotm` in the current user's Word Startup folder. It never changes `Normal.dotm`, Word macro-security policy, or machine-wide settings. Before replacement it saves a backup and retains the newest three.

Build both release files after refreshing the installed Startup template:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-installer.ps1
```

Publisher: MasseysLab. Software and documentation: MIT License.

See the repository's `docs/Release_Notes_v0.9.1-alpha.md` for release highlights,
known Alpha limits, and verification details.

## Publishing order

1. Merge the verified release commit into `main` and confirm the Build and Test
   workflow passes.
2. Create and push tag `v0.9.1-alpha` at that exact commit.
3. Create a GitHub Release for the tag and attach `CleanupSuite.dotm` with that
   exact filename. The update manifest and manual-install path depend on the
   resulting `/releases/download/v0.9.1-alpha/CleanupSuite.dotm` URL.
4. Also attach `CleanupSuiteForWord-Setup.exe` and the versioned user-manual PDF
   so the release page is self-contained.
5. Verify the stable Setup link, `latest-release.ini`, template download, and a
   manual Check for Updates from Word after publication.
