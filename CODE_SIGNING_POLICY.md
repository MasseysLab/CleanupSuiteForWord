# CleanupSuite Code Signing Policy

Free code signing provided by [SignPath.io](https://signpath.io/), certificate by
[SignPath Foundation](https://signpath.org/), is CleanupSuite's preferred
no-cost public-trust signing path, subject to SignPath Foundation accepting the
project.

## What a CleanupSuite Signature Means

A signed CleanupSuite release must:

- be built from the public
  [MasseysLab/CleanupSuiteForWord](https://github.com/MasseysLab/CleanupSuiteForWord)
  repository by the approved automated release pipeline;
- pass the repository tests, VBA assembly checks, engine contract tests, matched
  installer-package checks, and security scans;
- use consistent CleanupSuite product and version metadata;
- receive explicit signing approval after the release evidence is reviewed; and
- publish the signed artifacts and their SHA-256 values together.

The self-contained engine is signed before its hash is placed in the matched
installation manifest. The completed Setup is signed afterward. CleanupSuite
then validates both the signature and the exact manifest-bound package.

## Project Roles

- **Authors, committers, and reviewers:** the
  [MasseysLab repository owner](https://github.com/MasseysLab) and explicitly
  authorized CleanupSuite collaborators.
- **Signing approver:** the
  [MasseysLab repository owner](https://github.com/MasseysLab).

Repository and signing access require multi-factor authentication. Signing
approval is separate from authorship of the release change.

## Privacy and System Changes

CleanupSuite's network behavior and local document processing are described in
the [privacy policy](PRIVACY.md). Setup announces maintenance actions, installs
per user, preserves user settings, supplies repair and uninstall choices, and
does not weaken Word or Windows security settings.

## Unsigned Releases

An unsigned build must be clearly identified as unsigned and must publish its
SHA-256. Users should follow the
[unsigned-install safety guide](docs/Unsigned_Installation_Guide.md). The guide
does not ask users to disable SmartScreen, Microsoft Defender, Word macro
security, or the Trust Center.

Unsigned artifacts do not carry the assurances described in the signed-release
section above, even when their hashes match the repository release.
