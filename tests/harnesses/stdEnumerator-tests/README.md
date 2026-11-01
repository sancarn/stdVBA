# stdEnumerator Excel regressions

Run `python tests/harnesses/stdEnumerator-tests/run.py` on Windows with Excel and pywin32.
The runner uses a separate hidden Excel process and disposable workbook. It checks
actual bytes for CRLF before every import, monitors only its own Office windows,
dismisses dialogs as failures, and bounds worker execution. It discovers the
installed Office/VBE versions, respects managed AccessVBOM policy, and records,
restores and verifies the exact original registry values/types, including absence.

The existing stdEnumeratorTests suite runs with a lightweight assertion recorder
to avoid the original Test class's worksheet formatting dependency. The new probes
use a Python COM source with an instrumented IEnumVARIANT to verify:
- Zero constructor reads or cursor acquisitions; exactly 100 reads for First(100).
- Independent cursors for repeated, nested, procedural and native For Each traversals.
- Cursor release on early exit, failure, exhaustion and procedural refresh.
- Predicate callbacks interleaved with source reads during countBy.
- At most three live generated COM values while counting 3,000 values; no retained
  source values after the operation. First retains only its 100 returned objects.
- Source changes visible on later traversals, variant/object identity, dictionaries,
  Worksheets and a prefix of the full Excel Cells collection.

Verified on 64-bit Excel 16.0 / VBA 7.1. Windows 32-bit, legacy VBA and Mac runtime
paths remain untested. Mac retains eager enumeration. On Windows, Map, Filter,
Sort and explicit conversion operations still materialise their results; this
change provides streaming source traversal, not a fully lazy transformation pipeline.
