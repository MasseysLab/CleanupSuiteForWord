# CleanupSuite Analysis Engine

This directory contains the Task 5 foundation and Task 6 pilot for the future
C#/VBA hybrid. It is not installed or shipped as an official CleanupSuite
component.

`CleanupSuite.Engine`:

- implements Contract 1.0 from `contracts/hybrid/v1`;
- runs only as an on-demand command-line process;
- exposes `--version`, `--capabilities`, and `--analyze <job-directory>`;
- accepts only a lowercase-UUID job directly beneath the approved LocalAppData
  job root;
- reads only fixed protocol filenames;
- rejects reparse points, unknown JSON fields, mismatched hashes, invalid UTF-8,
  incompatible versions, and unsafe identity/path data;
- writes `result.json` atomically;
- returns no candidates on failure;
- produces content-free diagnostic logs;
- never references Word, a network API, a shell, or elevation.

Task 5 began with `contract-fixture/replace-literal`, which exists only to prove
deterministic candidates and fingerprints. Task 6 adds
`invisible-unicode-cleaner/selected-characters` as the first real pilot analyzer.
The source implementation is connected through the VBA bridge and has passed
automated Word Preview-data, Apply, single-Undo, stale-document rejection, and
large-document performance probes. It remains unavailable to ordinary users until
the matched self-contained engine is signed, installed, and visually validated
through the later release gates.

Run the standalone gate:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_hybrid_engine_tests.ps1
```

Prove the self-contained development packaging separately:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish_hybrid_engine_dev.ps1
```

The ordinary development build is framework-dependent. The publish script produces
an unsigned, self-contained, single-file `win-x64` proof under the ignored `build`
tree. The official product must use this self-contained packaging model: CleanupSuite
may use .NET internally, but users must not need to install, update, select, or
maintain a separate .NET runtime. Neither current output is a distributable hybrid
engine. Authenticode signing, reputation validation across clean machines, and
installer integration remain later release work.
