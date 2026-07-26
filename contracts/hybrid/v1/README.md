# CleanupSuite Hybrid Contract 1.0

## Boundary

Contract `1.0` defines an offline, file-based analysis exchange between
CleanupSuite's Word/VBA mediator and a future C# engine.

The engine:

- receives an immutable UTF-8 snapshot and analysis options;
- returns candidates, fingerprints, summaries, and sanitized diagnostics;
- never opens, automates, or edits Word;
- never applies a candidate;
- never decides that a stale candidate is close enough.

VBA:

- captures the Word story/scope without newline or character normalization;
- owns the user interface, Preview, Reading View, navigation, and highlighting;
- validates the full snapshot before accepting results into Preview;
- validates the full snapshot again before Apply;
- validates every candidate range and structural fingerprint before mutation;
- aborts the complete Apply before its first mutation if any validation fails;
- owns the single Word Undo record.

## Files and Transport

The initial transport is one owner-only job directory:

`%LOCALAPPDATA%\MasseysLab\CleanupSuite\Jobs\<job-id>\`

The only protocol filenames are:

- `request.json`
- `document.utf8.txt`
- optional `structure.json`
- `result.json`
- optional `cancel.request`
- optional `engine.log`

Messages may reference only these filenames, never absolute paths, parent
directories, URLs, or arbitrary filenames. Writers create `*.tmp` files in the
same directory, flush and close them, then atomically rename them to the protocol
filename. Reparse points and links are rejected. Jobs are deleted after VBA has
finished validation, except for a deliberate sanitized diagnostic retention mode.

The engine is launched directly, on demand, without `cmd.exe`, PowerShell, a
Windows service, elevation, or a network connection.

## Text and Offsets

`document.utf8.txt` is UTF-8 without a byte-order mark. It contains the exact Word
story text for the requested scope.

- Word paragraph marks remain `U+000D`.
- Word cell/end-of-row markers remain `U+0007`.
- No CR/LF conversion, Unicode normalization, trimming, or case conversion is
  permitted.
- All positions and lengths use UTF-16 code units, matching Word Range and BSTR
  behavior—not Unicode scalar-value or UTF-8 byte indexes.
- Candidate offsets are relative to the first UTF-16 code unit of the snapshot.
- `scope.startUtf16` and `scope.endUtf16` retain the absolute Word story positions
  needed for VBA to map a relative candidate back to Word.

This basis was verified directly in Word with `A😀B` followed by Word's required
paragraph mark. The returned story positions were `0–1` for `A`, `1–3` for the
supplementary-plane emoji, `3–4` for `B`, and `4–5` for the paragraph mark. Word
therefore treats the emoji as one `Characters` item but advances Range positions
by its two UTF-16 code units. The supplementary-Unicode fixture preserves this
regression case.

Snapshot creation and engine decoding must reject malformed/unpaired UTF-16
surrogates instead of silently replacing them. Such a snapshot returns
`invalid-snapshot-text` and produces no candidates.

## Snapshot and Candidate Revalidation

The snapshot identifier is the lowercase SHA-256 of the exact UTF-8 snapshot
bytes. Before Preview and before Apply, VBA recaptures the same scope and requires
an identical SHA-256.

Every applicable candidate also carries:

- the exact candidate-text hash;
- bounded prefix and suffix hashes and their UTF-16 lengths;
- an optional paragraph hash;
- a required structure hash for structural operations;
- an explicit fail-closed revalidation policy.

For Contract 1.0, all mutating candidates require:

- `requireWholeScopeSnapshot: true`
- `requireExactRange: true`
- `requireContext: true`
- `onMismatch: "abort-apply"`
- `allowRelocation: false`

Structural-destructive candidates additionally require
`requireStructure: true`.

VBA performs all checks for all selected candidates before opening the Undo record
or making the first document change. A mismatch discards the Apply attempt and
asks the user to run Preview again. Contract 1.0 has no approximate relocation
mechanism.

## Compatibility

There are three independent versions:

1. protocol/contract version;
2. engine product version;
3. tool-definition version.

The installed manifest binds the `.dotm`, engine executable, protocol, and hashes.
Before launch, VBA verifies the manifest, expected publisher, engine hash, and
version compatibility. The engine capability response must then agree with the
same protocol and tool-definition versions.

Contract 1.0 requires the same protocol major version. An engine may advertise a
minimum and maximum supported minor version. A tool-definition version must be
explicitly listed by the engine; it is never assumed compatible.

## Failure Behavior

Failures are fail-closed and never reuse partial output:

| Condition | Required VBA behavior |
| --- | --- |
| Engine missing, hash mismatch, bad signature, or version mismatch | Do not launch or Apply; identify the installation problem and recommend Repair. |
| Capability or tool-definition mismatch | Disable the engine-backed action; do not silently reinterpret the request. |
| Invalid request or malformed snapshot text | Make no candidates; show a concise analysis failure. |
| Timeout, cancellation, crash, or missing `result.json` | Terminate/close the job, ignore temporary or partial output, and make no document changes. |
| Invalid schema, wrong job/snapshot echo, oversized result, unknown operation, or unsafe path | Reject the complete result as a contract/security failure. |
| Scope or candidate fingerprint mismatch before Preview or Apply | Discard the result or abort Apply before its first mutation and require a new Preview. |

There is no silent fallback from an engine-backed mode to a behaviorally different
VBA analyzer. During incremental migration, an explicitly retained VBA mode may
remain available only when it is presented and tested as the same product behavior.
The standalone `v0.9.3-beta-vba-final` release remains the recovery choice for
environments that cannot run the engine.

## Trust and Privacy

- The engine runs as the current user and requires no elevation.
- It runs only for an explicit analysis request and exits afterward.
- It requires no network access and exposes no listening endpoint.
- The request uses an opaque document-session UUID; document paths and names are
  excluded.
- Logs exclude document text, candidate text, paths, and document names by
  default.
- The installer and runtime verify SHA-256 hashes.
- Authenticode signing under the MasseysLab publisher identity is required before
  the official hybrid Beta is distributed.
- The final VBA-only release remains available for environments that prohibit
  companion executables.

## Contract Files

- `protocol.json`: fixed filenames, compatibility rules, exit codes, and limits.
- `operation-vocabulary.json`: the operations the engine may propose.
- `request.schema.json`: VBA-to-engine analysis request.
- `response.schema.json`: engine-to-VBA result and candidate record.
- `capabilities.schema.json`: engine capability handshake.
- `installation-manifest.schema.json`: installed component/version/hash binding.
- `fixtures/`: deterministic valid examples used by automated tests.

Estimated combined difficulty: **42–57%**.
Estimated professional desirability: **97–100%**.
