# Testing stdVBA

Use [the shared runner](../Tools/Testing/README.md) for automated execution. `manifest.json` is the explicit inventory: it lists dependencies, independent cases, supported hosts, requirements and default inclusion. Root `*Tests.bas` modules remain in their established locations so existing workbook imports keep working. Test-only probes, recorders, fixture servers and specialized harnesses live in `harnesses/`; shared support lives in `support/`; Python runner/build tests live in `tooling/`, and native ABI tests in `native/`.

## Adding or changing a test

1. Keep behavioural assertions outside production classes. Put tests for a class in its existing `tests/std<Class>Tests.bas` module, or a new module in this folder. Preserve documentation examples in production source.
2. Prefer small public, parameterless test procedures with stable case IDs. Register each procedure explicitly in `manifest.json`; arbitrary public helpers are not discovered as tests. Make every case establish its own state and clean up its resources. Keep `testAll` wrappers for existing manual callers where useful.
3. List only required production sources and test helpers. Test framework dependencies belong to scratch projects, never to distributed production classes.
4. Use `Test.Assert label, condition`. The shared reporter counts failures without raising, so legacy `On Error Resume Next` cannot swallow failed assertions. Unexpected errors reaching the bridge fail the case; zero assertions also fail. Existing returned PASS/count/Boolean contracts have explicit adapters. Avoid new broad error suppression: it can still conceal an unexpected error before any assertion.
5. Mark host, platform and fixture requirements honestly. Native crash probes run in isolated processes and remain opt-in. A crash is a failure, never an expected pass. Keep benchmarks separate from correctness cases.
6. Check actual CRLF bytes before every import, including generated helpers. Preserve encodings. Run the selected case, its class suite, and the appropriate aggregate checks.

## Private internals and existing suites

Lambda VM tests now live in `support/stdLambdaVM.cls.inc`. The runner injects this fragment into a temporary stdLambda copy and converts Debug.Assert into counted assertions. The production class has no test dependency. Use `--suite stdLambda.vm`; the former production `protRunVMTests` test entry point has been removed.

Date cache tests similarly instrument a disposable source copy. WIP classes remain experimental; existing compile or behaviour failures are reported, not patched around by the runner.

Some legacy suites remain aggregate cases; split those as they are maintained. Callback has ModuleCalls, Bind and PointerCalls; Regex has Contract, NativeEngine, CaptureArraysAndLists and MatchAllCol. Root `Main.bas` and the worksheet/Immediate Window reporters in `lib/` are legacy manual-workbook tools. They are not the authoritative complete inventory. New automated suites use `support/Test.cls`.

Access host automation is not implemented. The old Clipboard suite requires a workbook-specific worksheet builder, contains hard-coded paths and documents known failures; it is explicitly unavailable in the automated inventory. The old WebSocket module is commented-out placeholder code; use `stdWebSocket.local`. The historical Regex thunk comparison also requires removed experimental CreateVBA/CreateNative APIs and is explicitly unavailable; current Regex cases cover the distributed class. These are visible skips under `--all`, not passing coverage. Existing WIP prototype classes can contain exploratory test methods; they are not production regression coverage.

Run the runner's own contract tests without Office with `python -B tests/tooling/test_runner.py`.
