# Shared test runner

Run commands from the repository root. Python 3.9+ is needed for discovery and tooling tests. Office suites require Windows, desktop Office and pywin32 (`python -m pip install pywin32`). Image tests additionally require Pillow; native suites list their prerequisites in their own harness documentation.

```powershell
python Tools/Testing/run.py --list
python Tools/Testing/run.py
python Tools/Testing/run.py --all --json test-results.json
python Tools/Testing/run.py --suite stdCallback
python Tools/Testing/run.py --suite stdCallback --test Bind
python Tools/Testing/run.py --suite stdRegex --test MatchAllCol --host Word
python Tools/Testing/run.py --suite stdLambda.vm --test Instructions
python Tools/Testing/run.py --suite tooling.regexBuild
python Tools/Testing/run.py --suite native.regexThunk --clang C:/path/to/llvm-mingw/bin/clang.exe
```

Suite and case IDs are case insensitive. Repeat `--suite` to choose multiple suites. `--test` selects an exact case ID, requires a suite, and fails on an unknown name. `--list --suite stdRegex` shows the available cases. Selection is by independently callable case, rather than filtering assertion output after executing an entire suite.

The default set runs unattended regressions. Legacy `Test.FullTest` branches are disabled unless `--full` is supplied; those branches can require UI or lengthy measurements. `--all` selects suites and does not implicitly enable these legacy branches. `--all` selects every registered suite, including UI, external-network, clipboard, native-call and experimental tests. Read `--list` requirements before using it. Host mismatches and unavailable legacy fixtures are explicitly reported as skipped. An explicitly selected skipped suite exits unsuccessfully; an aggregate run fails for any failure or if nothing passed. JSON output records each case, host, duration, status and available assertion details. Missing dependencies, imports, compile errors, dialogs, timeouts and crashes fail a case; remaining cases continue.

Each ordinary VBA case runs in a fresh disposable Office project. The supervisor records the real host/VBE versions, checks VBA project access, respects policy, temporarily adjusts AccessVBOM only if needed, selects Break on Unhandled Errors and restores the exact original registry values/types/absence in cleanup. It watches only owned Office windows for dialogs and bounds each worker to 90 seconds. Shared runner invocations refuse overlapping Office settings transactions. Specialized adapters retain their own process isolation and fixtures, within the shared settings transaction; their aggregate execution is bounded to ten minutes. Do not terminate the supervisor forcibly while it is restoring settings.

The former `Tools/*-tests` folders have moved under `tests/harnesses`; use this shared command instead. The Regex test command retains its generated-source check before calling the shared runner. Native builders remain build tools; regression test bodies live under `tests/`.
