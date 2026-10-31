# Agent guidance for stdVBA

These instructions apply throughout this repository. Read [contribution.md](contribution.md) before changing code, and follow the existing style of the affected class/module.

## Style and dependencies

- Keep classes usable by importing a single class wherever possible. Avoid introducing dependencies that force users to import a chain of additional classes or modules. The contribution guide permits shared generic interfaces such as `stdICallable`; keep other dependencies optional where possible.
- Follow the contribution guide's class-first design, `std`/`stdI` naming, predeclared instances, `Create` constructors, `protInit` initialisers, and `TThis` state conventions. Keep code usable across Office applications where feasible.
- Record breaking changes in `changelog.md`; also document relevant features and bug fixes as described in the contribution guide.

## Line endings: check before importing

- All text files created or edited must use CRLF (`\r\n`) line endings, with no lone LF or lone CR. Preserve the file's encoding. Do not convert binary files such as workbooks or `.frx` resources.
- Before importing any VBA source into the VBE, inspect the actual bytes of every file to be imported, including generated modules and test helpers. Normalize text to CRLF before import; do not wait for a test failure or CI to fix it.
- The current `.gitattributes` defaults most text files to LF and disables normalization for `.bas`, `.cls`, and `.frm`. Do not assume checkout, an editor, a generator, or Git has produced CRLF. Recheck after tools rewrite files and before finishing changes.
- CI's `.github/workflows/enforce_crlf.yml` covers `.bas`, `.frm`, and `.cls`; it does not replace the local pre-import check. Avoid unrelated repository-wide line-ending changes.

## Office test environment and VBA project access

- Prefer a separate, hidden Office instance and disposable test documents/workbooks. Keep the user's existing Office session and documents undisturbed.
- Programmatic access to the VBA project object model may be disabled. Before importing modules, check access for the actual host application and installed Office version (Excel, Word, PowerPoint, etc.).
- If access is disabled, locate that application's version-specific Trust Center registry setting (`AccessVBOM`) and temporarily enable it for testing as needed. Determine the correct application/version/key rather than assuming Excel or a fixed Office version. Record whether the value existed, its type, and its original value before changing it. Respect managed policy restrictions.
- Apply the setting before starting the test instance, restarting only the agent-owned instance if necessary. Restore the exact original registry state in cleanup, including removing a value that was originally absent. Do not leave VBA project access enabled as a side effect of testing.

## Error trapping and disruptive dialogs

- Before tests, record the VBE's current **Tools > Options > General > Error Trapping** setting. Options include **Break on All Errors**, **Break in Class Module**, and **Break on Unhandled Errors**.
- Use **Break on Unhandled Errors** for tests so errors intentionally handled by VBA do not interrupt execution. The other settings can cause breakpoints and error dialogs even when code contains an error handler.
- Restore the exact original error-trapping setting in cleanup, including on failure. Do not assume the original was **Break in Class Module** or leave the user's debugging preference changed.
- Imports, compilation, and test execution can show modal compile/error dialogs and bring Office or the VBE to the foreground, even when the instance is hidden and `DisplayAlerts` is disabled. Prepare dialog detection/dismissal before starting potentially blocking automation; a modal dialog can block a synchronous COM call.
- Monitor the agent-owned Office/VBE windows during import, compilation, and tests. If a compile/error dialog appears, dismiss it immediately and minimize the activated VBE and host windows as fast as possible to limit disruption to the user's current work. Capture the error text if possible without delaying dismissal. Target only the test instance; avoid blind keystrokes into the user's active window.
- Treat the dismissed error as a failed import/compile/test, diagnose and fix it, and rerun the affected checks. Dismissing a dialog is not a passing test. Use bounded waits so automation cannot hang indefinitely.
- Use failure-safe cleanup to restore settings and close only agent-owned test documents and instances. Report any setting that could not be restored and any checks that could not run.

## Preserve generated documentation

- Documentation comments are structured input to `Tools/VBDocsGen`, not disposable prose. Preserve descriptions, tagged comments, examples, remarks, and their association with the correct member. Avoid unrelated rewrites or deletions.
- When changing a signature or behavior, update the corresponding documentation in the same change: parameter names/types/defaults, return values, examples, and other affected tags. Add tags for new parameters and members; remove documentation only when the underlying API or documented behavior is intentionally removed.
- The generated artifact is **`docs.json`**. `.github/workflows/vba_docs_gen.yml` currently runs on pushes and generates it using `node ./Tools/VBDocsGen/main.js`.
- For changes to documented VBA code, run that generator and compare the resulting `docs.json` against the baseline before opening a PR. Check for unexpected removed classes, members, parameters, tags, descriptions, and examples, as well as stale signatures. Investigate and fix unintended losses in the source comments or generator input; do not accept a large deletion merely because generation succeeded.
- Keep generated changes limited to the intended API/documentation changes, verify their line endings, and report documentation generation or comparison failures rather than silently skipping them.
