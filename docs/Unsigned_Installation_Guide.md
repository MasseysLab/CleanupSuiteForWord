# Installing an Unsigned CleanupSuite Release Safely

CleanupSuite is free, open-source software. A release without an Authenticode
signature can still be installed, but Windows cannot use a certificate to prove
who published that downloaded file. These steps preserve Windows protections
while letting the user make an informed, one-time choice.

Use this guide for the current unsigned **0.9.3 Stable VBA-only Setup** or a
future release whose official release page explicitly labels Setup **unsigned**
and publishes its Setup SHA-256. The current `0.9.5-beta` hybrid candidate still
requires valid signatures and will refuse an unsigned production installation
while the free SignPath application is pursued.

## Before Running Setup

1. Download `CleanupSuiteForWord-Setup.exe` only from the official
   [MasseysLab/CleanupSuiteForWord GitHub repository](https://github.com/MasseysLab/CleanupSuiteForWord/releases).
2. Close every Microsoft Word window.
3. In File Explorer, right-click the downloaded Setup and select **Scan with
   Microsoft Defender**. On some Windows versions, first select **Show more
   options**.
4. Verify the file's SHA-256. Open PowerShell and run:

   ```powershell
   Get-FileHash "$env:USERPROFILE\Downloads\CleanupSuiteForWord-Setup.exe" `
     -Algorithm SHA256
   ```

5. Compare the complete 64-character result with the Setup SHA-256 published in
   the same CleanupSuite release. Do not continue if even one character differs.

## The Windows SmartScreen Message

If Windows shows **Windows protected your PC**:

1. Confirm that the download came from the official repository, its SHA-256
   matches, and Microsoft Defender found no threat.
2. Select **More info**.
3. Confirm that the displayed app name is
   `CleanupSuiteForWord-Setup.exe`. An unsigned release may display **Unknown
   publisher**; the verified hash is what ties this exact file to the release.
4. Select **Run anyway** only after all checks pass.
5. In CleanupSuite Setup, review the detected state and choose **Install**,
   **Update**, or **Repair/Reinstall** as appropriate.

If **Run anyway** is unavailable, Windows or an organization administrator has
chosen a stricter policy. Do not disable or bypass that policy. Use an approved
computer, ask the administrator, or use a compatible standalone VBA-only
release.

## Standalone VBA-Only Alternative

Use this method only when the release is explicitly labeled **standalone** or
**VBA-only**:

1. Download `CleanupSuite.dotm` from the official GitHub release.
2. Verify its published SHA-256 and scan it with Microsoft Defender.
3. Close Word.
4. If File Explorer shows an **Unblock** checkbox in the file's **Properties**,
   select **Unblock**, then **Apply** and **OK**.
5. Copy `CleanupSuite.dotm` to:

   ```text
   %APPDATA%\Microsoft\Word\STARTUP
   ```

6. Reopen Word and select **CleanupSuite** on the ribbon.

Do not use this template-only method for a hybrid release. Hybrid CleanupSuite
versions require Setup to install the matched template, analysis engine,
protocol, rules, and installation manifest together.

## Protections That Should Stay Enabled

Installing CleanupSuite never requires you to:

- turn off Microsoft Defender or SmartScreen;
- weaken Word macro-security settings;
- change the Trust Center;
- add a broad trusted location;
- edit `Normal.dotm`; or
- run Setup as administrator.

Only override a warning for a file you intentionally downloaded from the
official repository, verified by SHA-256, and scanned. Microsoft similarly
recommends unblocking only files from trusted sources after checking the file:
[Microsoft Attachment Manager guidance](https://support.microsoft.com/en-us/windows/security/information-about-the-attachment-manager-in-microsoft-windows).

## Uninstall

Setup installs per user and provides **Uninstall** when run against an installed
copy. For a standalone VBA-only installation, close Word and remove:

```text
%APPDATA%\Microsoft\Word\STARTUP\CleanupSuite.dotm
```
