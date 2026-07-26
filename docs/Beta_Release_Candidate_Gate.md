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

- A trusted Authenticode code-signing certificate with its private key must be
  available to the build account.
- The self-contained engine and completed Setup must both report a valid,
  timestamped Authenticode signature.
- The signed candidate must pass the complete automated gate and Microsoft
  Defender scanning.
- Install, update, repair, uninstall, Word loading, Preview, Apply, and Undo must
  pass on a clean Windows/Word test account or machine.
- Only the candidate that passed those checks may be copied to `release` and
  published.

## Signed Candidate Command

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\build-installer.ps1 `
  -TemplatePath <fresh-synchronized-dotm> `
  -EnginePath <self-contained-engine> `
  -SignBuild `
  -SigningCertificateThumbprint <thumbprint>
```

This writes to `build\installer-signed-candidate` by default. It does not modify
the official release. The output includes the signed Setup and a
`candidate-payload` copy of the exact signed engine and matched package for
independent signature, hash, and security inspection.

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
