# stdEnumerator Excel regressions

Run `python Tools/stdEnumerator-tests/run.py` on Windows with Excel and pywin32.
The runner uses a separate hidden Excel process and disposable workbook. It checks
actual bytes for CRLF before every import, monitors only its own Office windows,
dismisses dialogs as failures, and bounds worker execution. It discovers the
installed Office/VBE versions, respects managed AccessVBOM policy, and records,
restores and verifies the exact original registry values/types, including absence.

The existing stdEnumeratorTests suite runs with a lightweight assertion recorder
to avoid the original Test class's worksheet formatting dependency. The new probes
use a Python COM source with an instrumented IEnumVARIANT to verify zero constructor
reads, exact prefix reads, replay, delayed errors and terminal exhaustion. They
also cover interleaved procedural reads, refresh, native For Each, dictionaries,
mixed variants/object identity and a prefix of the full Excel Cells collection.

Verified on 64-bit Excel 16.0 / VBA 7.1. The Windows 32-bit and legacy VBA paths
are included but have not been executed. Mac retains its existing eager fallback;
Mac runtime testing is unavailable here. Native For Each still materialises the
source through AsCollection. Cached values take memory proportional to consumption;
unread source mutations and source errors are observed during consumption.
