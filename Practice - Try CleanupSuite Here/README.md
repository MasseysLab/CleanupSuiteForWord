<!-- SPDX-License-Identifier: MIT -->
<!-- Copyright (c) 2026 MasseysLab -->

# Practice CleanupSuite Safely

This folder contains a macro-enabled practice copy of CleanupSuite and small
test documents for learning what each tool changes before using it on important
work.

## Quick practice

1. Close other Word documents you do not need.
2. Open `CleanupSuite.docm` and enable its macros if you trust this copy.
3. Open the matching document under `Testing Files`.
4. Select **CleanupSuite** on Word's ribbon, choose the named tool, and run its
   Preview.
5. Compare the tan sample with the green expected result before applying.
6. Close the test document without saving when you want to restore the original
   sample.

The global per-user installer remains the preferred normal installation method.
See the project [README](../README.md) for installation instructions and the
[CleanupSuite User Manual](../documents/CleanupSuite_User_Manual.pdf) for the
complete tool guide.

## Developer note

`VBA_Cleanup_tool.txt` is the generated pasteable VBA bundle. Its readable
source lives under the repository's `src` directory and should be rebuilt with
`python assemble.py`; do not edit the generated bundle directly.

Software and documentation are licensed under the [MIT License](../LICENSE).
