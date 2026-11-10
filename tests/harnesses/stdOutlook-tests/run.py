"""Run VBA in a disposable Excel process using stdVBA's safeguarded test host.

Requires pywin32, Excel and classic Outlook; uses this checkout's shared Office host.
The production class has no dependency on this test-only tooling.
"""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import sys
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[3]
HERE = Path(__file__).resolve().parent

# Append test entry points to a TEMPORARY copy, never the distributable class.
PROBES = '''
Friend Function TestDecode(ByVal pointer As LongPtr) As Variant
  TestDecode = DecodeValue(pointer)
End Function
Friend Function TestProperties(ByVal pointer As LongPtr, ByVal count As Long, ByVal columns As Variant) As Collection
  Set TestProperties = DecodeProperties(pointer, count, columns)
End Function
Friend Function TestTable(ByVal pointer As LongPtr, ByVal limit As Long, ByRef capped As Boolean) As Collection
  Set TestTable = ReadTable(pointer, Array(Array("Number", &H30000003), Array("Text", &H3001001F)), limit, capped)
End Function
Friend Sub TestRestriction(ByVal pointer As LongPtr, ByVal query As String)
  RestrictName pointer, query
End Sub
Friend Sub TestRelease(ByRef pointer As LongPtr)
  ReleaseObject pointer
End Sub
Friend Function TestEntryID(ByVal value As String) As String
  Dim bytes() As Byte
  bytes = DecodeEntryID(value)
  TestEntryID = HexBytes(VarPtr(bytes(0)), UBound(bytes) + 1)
End Function
Friend Function TestPerson(ByVal pointer As LongPtr) As Collection
  Set TestPerson = ReadPerson(pointer)
End Function
Friend Sub TestTransport(ByVal session As LongPtr, ByVal book As LongPtr, ByVal store As LongPtr, Optional ByVal writable As Boolean = True)
  This.Session = session
  This.AddressBook = book
  This.Store = store
  This.StoreWritable = writable
End Sub
Friend Function TestAppointmentTag(ByVal folder As LongPtr, ByVal lid As Long, ByVal kind As Long) As Long
  TestAppointmentTag = AppointmentTag(folder, lid, kind)
End Function
Friend Sub TestCalendarRestriction(ByVal table As LongPtr, ByVal query As String, ByVal columns As Variant)
  RestrictCalendar table, query, columns
End Sub
'''


def crlf(data: bytes) -> bytes:
    return data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stdvba", type=Path, default=ROOT, help="Path to stdVBA checkout (defaults to this repository)")
    parser.add_argument("--live", action="store_true", help="Also read the default Outlook profile; no personal data is logged")
    args = parser.parse_args()
    host_file = args.stdvba.resolve() / "Tools/Testing/office.py"
    if not host_file.is_file():
        parser.error("stdVBA Tools/Testing/office.py was not found")
    with tempfile.TemporaryDirectory(prefix="stdOutlook-tests-") as temp:
        scratch = Path(temp)
        # Keep stdVBA's checkout unchanged. Add an explicit VBE compilation step
        # to its temporary runner; its existing watchdog treats any dialog as failure.
        host_source = host_file.read_text(encoding="utf-8")
        # ShowWindow can block behind a busy Office GUI thread, defeating the
        # parent deadline. The async form keeps the watchdog independent.
        host_source = host_source.replace("USER.ShowWindow", "USER.ShowWindowAsync")
        host_source = host_source.replace(
            '            macro = "UnifiedTestRunner.RunCase"',
            '            print("Imported modules:", [(c.Name, c.Type, c.CodeModule.CountOfLines) for c in project.VBComponents], flush=True)\n'
            '            compile_command = app.VBE.CommandBars.FindControl(1, 578)\n'
            '            if compile_command is not None and compile_command.Enabled:\n'
            '                compile_command.Execute()\n'
            '                print("Compile command returned", flush=True)\n'
            '                pane = app.VBE.ActiveCodePane\n'
            '                if pane is not None:\n'
            '                    selection = pane.GetSelection()\n'
            '                    print("VBE location:", pane.CodeModule.Name, selection, pane.CodeModule.Lines(selection[0], 1), flush=True)\n'
            '            macro = "UnifiedTestRunner.RunCase"',
        )
        host_source = host_source.replace(
            '            if dialogs:\n                raise RuntimeError',
            '            if dialogs:\n                time.sleep(0.25)  # Let the worker report the selected compile-error line.\n                raise RuntimeError',
        )
        staged_host = scratch / "office_host.py"
        staged_host.write_bytes(crlf(host_source.encode("utf-8")))
        spec = importlib.util.spec_from_file_location("stdvba_office", staged_host)
        office = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(office)
        source = (ROOT / "src/stdOutlook.cls").read_bytes()
        # Reject bad repository bytes before doing ANY VBA import.
        files = [ROOT / "src/stdOutlook.cls", ROOT / "tests/stdOutlookTests.bas", HERE / "stdOutlookSendCalendarTests.bas", HERE / "stdOutlookSyncFake.cls", ROOT / "tests/support/Test.cls"]
        for path in files:
            data = path.read_bytes()
            if data != crlf(data):
                raise RuntimeError(f"CRLF required before import: {path}")
        staged_class = scratch / "stdOutlook.cls"
        # Replace acquisition only in the staged class. All production sync logic
        # runs against a late-bound fake, even with --live. Fail closed on drift.
        for original, replacement in (
            (b'GetObject(, "Outlook.Application")', b'stdOutlookSendCalendarTests.GetRunningOutlookFake()'),
            (b'CreateObject("Outlook.Application")', b'stdOutlookSendCalendarTests.StartOutlookFake()'),
        ):
            if source.count(original) != 1:
                raise RuntimeError("Outlook acquisition changed; update safe test substitution")
            source = source.replace(original, replacement)
        staged_class.write_bytes(crlf(source + PROBES.encode("ascii")))
        staged_tests = scratch / "UnifiedTestRunner.bas"
        test_source = files[1].read_bytes().replace(b'Attribute VB_Name = "stdOutlookTests"', b'Attribute VB_Name = "UnifiedTestRunner"').replace(b"#Const LIVE_TEST = False", b"#Const LIVE_TEST = " + str(args.live).encode("ascii"))
        staged_tests.write_bytes(crlf(test_source))
        staged_extra = scratch / "stdOutlookSendCalendarTests.bas"
        staged_extra.write_bytes(crlf(files[2].read_bytes()))
        staged_outlook = scratch / "stdOutlookSyncFake.cls"
        staged_outlook.write_bytes(crlf(files[3].read_bytes()))
        staged_reporter = scratch / "Test.cls"
        staged_reporter.write_bytes(crlf(files[4].read_bytes()))
        # stdVBA's host probes the installed Office/VBE versions, checks policy,
        # records/restores AccessVBOM and both error-trapping registry values,
        # owns a separate Excel PID, monitors/dismisses only its modal dialogs,
        # imposes a 90-second deadline, and locks against overlapping test runs.
        progress = scratch / "progress.txt"
        old_progress = os.environ.get("STDOUTLOOK_TEST_PROGRESS")
        os.environ["STDOUTLOOK_TEST_PROGRESS"] = str(progress)
        try:
            office.run({"kind": "vba", "files": [str(staged_class), str(staged_tests), str(staged_extra), str(staged_outlook), str(staged_reporter)]}, "Excel", temp)
        finally:
            if old_progress is None:
                os.environ.pop("STDOUTLOOK_TEST_PROGRESS", None)
            else:
                os.environ["STDOUTLOOK_TEST_PROGRESS"] = old_progress
            if progress.exists():
                print(progress.read_text())
        result = json.loads((scratch / "result.json").read_text(encoding="utf-8"))
        print(json.dumps(result, indent=2))
        if result.get("status") != "passed" or result.get("assertions", 0) <= 0:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
