Run `powershell -NoProfile -File tests/harnesses/stdDate-test/run-tests.ps1` from this worktree.

The compatibility command now delegates to `python Tools/Testing/run.py --suite stdDate`. The shared lifecycle uses a disposable hidden Excel instance, monitors modal dialogs, checks CRLF before import, and restores exact VBA project-access and error-trapping settings. A temporary class copy reports cache dictionary counts; it is discarded after the run. Production source is not instrumented. Tests cover weekday defaults/names and regex cache behaviour. stdDate remains experimental.

2026-10-31: execution blocked at VBProject access by the existing trust setting. No VBA assertions ran and compilation was not verified. The WIP class also contains a pre-existing Exit Function inside the isLeap Property Get; this work does not repair unrelated WIP functionality.
