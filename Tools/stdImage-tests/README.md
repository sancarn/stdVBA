# stdImage regressions

Run from the repository root on Windows with Microsoft Office, Python, pywin32 and Pillow installed:

```powershell
python Tools/stdImage-tests/run.py
python Tools/stdImage-tests/run.py --hosts Excel
```

The runner creates disposable, hidden Excel and Word instances and refuses to use a process that already existed. It imports only stdImage and the standalone stdImageTests module, checking actual CRLF bytes before each import. An independent parent watches the owned process for modal dialogs, captures their text, dismisses them and minimises its Office/VBE windows; any dialog, timeout or failed assertion fails the run.

Before importing, a probe determines the actual host and VBE versions and tests VBA project access. When necessary, the runner temporarily enables that host's AccessVBOM value unless managed policy denies access. It records the existence, registry type and value of AccessVBOM and both VBE error-trapping values, uses Break on Unhandled Errors, and restores the exact original values (including absence) in finally blocks. The original clipboard is retained through OLE and restored after the run. Close and forced cleanup target only owned Office processes.

Fixtures are generated in a disposable directory: BMP, PNG, GIF, JPEG, TIFF, WMF, EMF, bitmap-backed ICO, transparent PNG and a JPEG quality pattern. A local HTTP fixture serves image bytes and HTTP errors. The tests cover all output formats and MIME types, defaults, binary/file/Data URL round-trips, pixels and dimensions, nonzero byte-array bounds, malformed inputs, transparency flattening, JPEG/TIFF parameters, independent picture and clipboard lifetimes, temporary-file cleanup and GDI handle stability after successful and failed conversions. Excel additionally exercises worksheet embedding and picture/fill helpers.

## Recorded validation

- 64-bit Office 16: 237 assertions passed in Excel and 235 in Word.
- 32-bit Office and legacy VBA are not installed here; those runtime checks remain unverified.
- Core loading rasterises the first frame against the selected background (black by default). Animation, metadata, vector content and alpha are not retained. WebP is outside this change and tracked in issue #178.
- CreateFromUrl retains default WinHTTP proxy behaviour; callers needing custom proxy/authentication/timeout handling can use stdHTTP and CreateFromBinary.

After changing documented source, run `node Tools/VBDocsGen/main.js`. Compare docs.json by class/member identity to catch lost tags, examples or signatures. The generator's directory order differs by platform; preserve the baseline class order to avoid unrelated ordering changes. Normalise generated docs.json to CRLF.
