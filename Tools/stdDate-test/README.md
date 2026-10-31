Run `powershell -NoProfile -File Tools/stdDate-test/run-tests.ps1` from this worktree.

Uses a new hidden Excel application and scratch workbook; closes only these objects and changes no security settings. Requires existing trust for VBA project access. A temporary imported stdDate copy reports actual static dictionary counts to the runner; production source/API is not instrumented. Tests cover US/UK defaults, all seven dates and explicit first days, abbreviated/full names, regex reuse/options, malformed errors and cache retention. The temporary class file is retained in the system temp directory for inspection.

2026-10-31: execution blocked at VBProject access by the existing trust setting. No VBA assertions ran and compilation was not verified. The WIP class also contains a pre-existing Exit Function inside the isLeap Property Get; this work does not repair unrelated WIP functionality.
