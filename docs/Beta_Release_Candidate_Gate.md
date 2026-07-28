# CleanupSuite 0.9.5-beta Release-Candidate Gate

The official hybrid Beta must not replace the stable release until every gate
below passes. Candidate creation and official publication are separate
operations so a failed signature, scan, or validation cannot disturb the
published package. Publication also compiles and signs in staging, captures the
three existing release artifacts, and restores them if a replacement fails.

## Passed Gates

- Hybrid Contract 1.0 and exact capability validation.
- Self-contained engine; users do not install or maintain .NET separately.
- Word-to-engine Unicode pilot with stale-document rejection and one-step Undo.
- Minimal-change highlighting and review paths for unhighlighted findings.
- One state-aware Setup for Install, Update, Repair/Reinstall, and Uninstall.
- Matched offline payload with version, path, byte-length, and SHA-256 binding.
- Transactional maintenance rollback, newest-three backups, and settings
  preservation.
- Isolated absent, legacy, older, current, damaged, repair, update, rollback,
  and uninstall matrix.
- Production installed-manifest Word smoke and live Setup layout review.
- Stable `release` artifacts remain untouched.

## Blocking Gates

- A trusted public-signing path must be available. The preferred no-cost route is
  sponsored signing through
  [SignPath Foundation](https://signpath.org/), subject to application approval.
  A conventional Authenticode certificate remains an optional paid route.
- The self-contained engine and completed Setup must both report a valid,
  timestamped Authenticode signature.
- The signed candidate must pass the complete automated gate and Microsoft
  Defender scanning.
- Install, update, repair, uninstall, Word loading, Preview, Apply, and Undo must
  pass on a clean Windows/Word test account or machine.
- Only the candidate that passed those checks may be copied to `release` and
  published.

The [free open-source signing plan](Free_Open_Source_Code_Signing_Plan.md)
records the SignPath eligibility and integration work. If the application is
rejected, an explicitly unsigned release with published hashes and the
[unsigned-install safety guide](Unsigned_Installation_Guide.md) is a separate
distribution decision; it is not a successful completion of this signed gate.

## Local-Certificate Candidate Command

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-installer.ps1 `
  -TemplatePath <fresh-synchronized-dotm> `
  -EnginePath <self-contained-engine> `
  -SignBuild `
  -SigningCertificateThumbprint <thumbprint>
```

This optional paid/local-certificate path writes to
`build\installer-signed-candidate` by default. It does not modify the official
release. The output includes the signed Setup and a
`candidate-payload` copy of the exact signed engine and matched package for
independent signature, hash, and security inspection.

An accepted SignPath integration will perform the equivalent nested
engine-then-Setup signing through its repository-linked build pipeline rather
than a local certificate thumbprint.

Run the repeatable local candidate gate next:

```powershell
powershell -ExecutionPolicy Bypass -File `
  .\scripts\run_beta_release_candidate_gate.ps1
```

It rejects unknown, duplicate, missing, path-mismatched, length-mismatched, or
hash-mismatched components; requires valid Authenticode signatures on the engine
and Setup; requires Microsoft Defender and real-time protection; and scans those
exact signed binaries.

## Publication Command

Run this only after the signed candidate passes every blocking gate:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-installer.ps1 `
  -TemplatePath <validated-synchronized-dotm> `
  -EnginePath <validated-self-contained-engine> `
  -PublishRelease `
  -SigningCertificateThumbprint <thumbprint>
```

Estimated combined difficulty: **18–30% remaining after a certificate is
available**.

Estimated professional desirability: **99–100%**.
