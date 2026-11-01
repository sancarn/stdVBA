# Issue 85 object argument probes

Run `python tests/harnesses/stdCallback-tests/run.py` from the repository root with Windows,
Excel, Python and pywin32 installed. Each scenario uses a separate disposable
Excel instance. Do not combine the ByRef probes into the ordinary test suite:
passing an object value where a pointer to an object variable is required can
terminate Excel. The runner bounds each call, watches for modal dialogs, and
restores the original registry values for VBA error trapping and Excel VBA
project access. It prints those original values before testing.

Observed on 64-bit Excel 16.0.20618.20000:

- ByVal Object, inferred types: identity preserved.
- ByVal Object, explicit vbObject: identity preserved.
- ByRef Object, inferred types, ordinary object argument: COM call failed and
  the Excel instance disconnected.
- ByRef Object, explicit vbObject, ordinary object argument: same failure.
- ByRef Object, VarPtr(objectVariable), inferred pointer type: identity preserved.
- ByVal Variant, explicit vbVariant, containing an object: identity preserved.
- ByVal Collection, inferred types: identity preserved.
- ByVal Object through RunEx with an object stored in a Variant array: identity
  preserved.
- ByVal Object with Nothing: identity preserved.
- ByVal Object with Scripting.Dictionary: identity preserved.

The probes check object identity using `Is`, rather than only checking TypeName.
These are characterization probes, including deliberately invalid native calls;
a process failure is a reproduced failure, never a passing assertion. A successful
case must print `PASS identity=True` with no dialogs and exit zero.

This supports a ByRef calling-convention explanation for issue 85 on this host.
VBA defaults object parameters to ByRef; vbObject alone does not encode ByRef.
The original report does not contain a reproducer, so it is not possible to prove
that this was its exact scenario. No production implementation is changed.
32-bit Office, other Office hosts, object return values, ByRef reassignment and
explicit VT_BYREF type flags have not been tested.
