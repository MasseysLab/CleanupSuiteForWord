# CleanupSuite Privacy Policy

CleanupSuite processes Word documents locally on the user's Windows computer. It
does not provide accounts, advertising, analytics, telemetry, or cloud document
processing. CleanupSuite does not sell, share, or upload document content,
document names, document paths, user identity, or usage information.

## Local Document Processing

The VBA tools operate inside Microsoft Word. Hybrid analysis uses a
self-contained local engine installed with CleanupSuite. For an analysis, the
selected document text and request metadata are written to an owner-only
temporary job directory under the current user's local application-data folder.
The local engine reads that job, writes its result locally, and CleanupSuite
attempts to delete the job after completion or handled failure. An abrupt Word
or Windows termination can leave an owner-only job for local recovery or manual
deletion. The engine opens no network listener, runs no service, requires no
elevation, and does not transmit document content.

## Network Activity

When periodic update checks are enabled, CleanupSuite contacts only the official
GitHub-hosted release manifest at most once every seven days. The request
identifies the installed CleanupSuite version through its user-agent string. It
does not include document content, document names, document paths, or user
identity. Periodic checks can be disabled in the CleanupSuite launcher; the
manual **Check for Updates** control remains available.

CleanupSuite opens the official GitHub user manual or release/download page only
when the user selects the corresponding command or accepts an offered update
link. GitHub's handling of those web requests is governed by
[GitHub's Privacy Statement](https://docs.github.com/en/site-policy/privacy-policies/github-general-privacy-statement).

## Local Settings and Backups

CleanupSuite stores user preferences in the current user's Windows settings. The
installer may retain the newest three local template backups and preserves user
settings during maintenance and uninstall. CleanupSuite does not transmit those
settings or backups.

## Security Scope

CleanupSuite does not modify `Normal.dotm`, Word macro-security policy, the Trust
Center, trusted locations, or machine-wide registry settings. Installation and
operation are per user.

Questions or corrections can be submitted through the official
[CleanupSuite GitHub repository](https://github.com/MasseysLab/CleanupSuiteForWord/issues).
