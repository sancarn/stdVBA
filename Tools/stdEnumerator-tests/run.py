"""Run bounded lazy-enumeration regressions in disposable Excel. Requires pywin32.
The parent dismisses owned dialogs and restores exact Office/VBE registry values.
"""
import argparse
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import winreg

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
USER = ctypes.WinDLL("user32", use_last_error=True)
USER.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
USER.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
USER.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
USER.GetDlgCtrlID.argtypes = [wintypes.HWND]
USER.GetDlgCtrlID.restype = ctypes.c_int
USER.GetDlgItem.argtypes = [wintypes.HWND, ctypes.c_int]
USER.GetDlgItem.restype = wintypes.HWND
USER.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
USER.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
CALLBACK = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
USER.EnumWindows.argtypes = [CALLBACK, wintypes.LPARAM]
USER.EnumChildWindows.argtypes = [wintypes.HWND, CALLBACK, wintypes.LPARAM]


def read_value(path, name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            return winreg.QueryValueEx(key, name)
    except FileNotFoundError:
        return None


def restore_value(path, name, original):
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
        if original is None:
            try:
                winreg.DeleteValue(key, name)
            except FileNotFoundError:
                pass
        else:
            winreg.SetValueEx(key, name, 0, original[1], original[0])
    if read_value(path, name) != original:
        raise RuntimeError("Could not restore " + path + "\\" + name)


def worker(host, scratch, phase, url):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    import win32process
    existing_pids = set(win32process.EnumProcesses())
    app = doc = None
    pidpath = Path(scratch) / "office.pid"
    try:
        app = win32com.client.DispatchEx(host + ".Application")
        pid = wintypes.DWORD()
        if host == "Word":
            # Word exposes a document window handle without requiring VBE trust.
            doc = app.Documents.Add(Visible=False)
        hwnd = app.Hwnd if host == "Excel" else doc.ActiveWindow.Hwnd
        USER.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
        if pid.value in existing_pids:
            app = None
            raise RuntimeError("Office returned an existing process; refusing to use it")
        pidpath.write_bytes(str(pid.value).encode("ascii"))
        app.Visible = False
        app.DisplayAlerts = 0
        if host == "Excel":
            doc = app.Workbooks.Add()
        if phase == "probe":
            try:
                project = doc.VBProject
                access = project.VBComponents.Count >= 0
            except Exception:
                access = False
            info = {"office": str(app.Version), "vbe": str(app.VBE.Version) if access else None, "access": access}
            (Path(scratch) / "probe.json").write_bytes(json.dumps(info).encode("utf-8"))
        else:
            project = doc.VBProject
            count = project.VBComponents.Count
            zone = doc.Worksheets(1)
            project.VBComponents(zone.CodeName).Name = "TestZone"
            for relative in ("src/stdICallable.cls", "src/stdEnumerator.cls", "src/stdCallback.cls", "src/stdLambda.cls", "Tools/stdEnumerator-tests/Test.cls", "tests/stdEnumeratorTests.bas", "Tools/stdEnumerator-tests/Probe.bas"):
                raw = (ROOT / relative).read_bytes()
                if b"\n" in raw.replace(b"\r\n", b"") or b"\r" in raw.replace(b"\r\n", b""):
                    raise RuntimeError("CRLF required before import: " + relative)
                print("Importing " + relative, flush=True)
                project.VBComponents.Import(str(ROOT / relative))
                print("Imported " + relative, flush=True)
            # Running invokes the compiler; modal failures are caught by the parent.
            macro = "'" + doc.Name + "'!EnumeratorProbe.RunTests" if host == "Excel" else "EnumeratorProbe.RunTests"
            print("Running " + macro, flush=True)
            from win32com.server.util import wrap, ListEnumerator
            from win32com.server.exception import COMException
            class Cursor(ListEnumerator):
                def Next(self, count):
                    source.reads += 1
                    if source.fail_after and self.index >= source.fail_after:
                        raise COMException("Intentional source failure", scode=-2147467259)
                    return super().Next(count)
            class Source:
                _public_methods_ = ["Reads", "Configure"]
                def __init__(self):
                    self.reads = 0
                    self.fail_after = 0
                def _NewEnum(self):
                    return wrap(Cursor(list(range(1, 1001))), pythoncom.IID_IEnumVARIANT)
                def Reads(self):
                    return self.reads
                def Configure(self, fail_after):
                    self.reads = 0
                    self.fail_after = fail_after
            source = Source()
            suite_result = app.Run("'" + doc.Name + "'!EnumeratorProbe.RunSuite")
            print("Suite:", suite_result, flush=True)
            if not suite_result.startswith("PASS:"):
                raise RuntimeError(suite_result)
            result = app.Run(macro, wrap(source))
            print(host + ": " + str(result), flush=True)
            if not str(result).startswith("PASS:"):
                raise RuntimeError(str(result))
    except Exception:
        try:
            pane = app.VBE.ActiveCodePane
            selection = pane.GetSelection()
            print("VBE selection: " + repr(selection), flush=True)
            print("VBE source: " + str(pane.CodeModule.Lines(selection[0], 1)), flush=True)
        except Exception as location_error:
            print("Could not capture VBE location: " + str(location_error), flush=True)
        raise
    finally:
        try:
            if doc is not None:
                doc.Close(False if host == "Excel" else 0)
        finally:
            if app is not None:
                app.Quit()
            doc = app = None
            pythoncom.CoUninitialize()


def bounded_worker(host, scratch, phase, url):
    pidpath = Path(scratch) / "office.pid"
    if pidpath.exists():
        pidpath.unlink()
    exports = Path(scratch) / "exports"
    exports.mkdir(exist_ok=True)
    environment = dict(os.environ, TEMP=str(exports), TMP=str(exports))
    process = subprocess.Popen([sys.executable, __file__, "--worker", host, scratch, phase, url], env=environment)
    owned = None
    deadline = time.monotonic() + 90
    dialogs = []
    try:
        while process.poll() is None:
            if pidpath.exists():
                text = pidpath.read_text()
                if text:
                    owned = int(text)

            @CALLBACK
            def inspect(hwnd, unused):
                pid = wintypes.DWORD()
                USER.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if owned and pid.value == owned:
                    kind = ctypes.create_unicode_buffer(256)
                    USER.GetClassNameW(hwnd, kind, len(kind))
                    if kind.value == "#32770":
                        parts = []

                        @CALLBACK
                        def capture(child, ignored):
                            label = ctypes.create_unicode_buffer(1024)
                            USER.GetWindowTextW(child, label, len(label))
                            if label.value:
                                parts.append(label.value)
                            if label.value.replace("&", "") == "OK":
                                control_id = USER.GetDlgCtrlID(child)
                                USER.PostMessageW(hwnd, 0x0111, control_id, child)
                                USER.PostMessageW(child, 0x00F5, 0, 0)
                            return True

                        USER.EnumChildWindows(hwnd, capture, 0)
                        dialogs.append(" | ".join(parts))
                        print("Detected and dismissed test dialog: " + dialogs[-1], flush=True)
                        # Dismiss only this test instance's dialog; never send global keys.
                        button = USER.GetDlgItem(hwnd, 1)
                        if button:
                            USER.PostMessageW(hwnd, 0x0111, 1, button)
                            USER.PostMessageW(button, 0x00F5, 0, 0)
                            USER.PostMessageW(hwnd, 0x0100, 13, 0)
                            USER.PostMessageW(hwnd, 0x0101, 13, 0)
                        else:
                            USER.PostMessageW(hwnd, 0x0010, 0, 0)
                    USER.ShowWindow(hwnd, 6)
                return True

            USER.EnumWindows(inspect, 0)
            if dialogs:
                raise RuntimeError("Office dialog (failed check): " + repr(dialogs))
            if time.monotonic() > deadline:
                raise RuntimeError("Office worker timed out; dialogs=" + repr(dialogs))
            time.sleep(0.05)
        if dialogs:
            raise RuntimeError("Office dialog (failed check): " + repr(dialogs))
        if process.returncode:
            raise RuntimeError(host + " worker failed")
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=5)
        if owned:
            # Only the PID reported by our DispatchEx-created instance.
            subprocess.run(["taskkill", "/PID", str(owned), "/F"], capture_output=True)


def main():
    import win32com.client
    with tempfile.TemporaryDirectory(prefix="stdEnumerator-") as scratch:
        bounded_worker("Excel", scratch, "probe", "")
        info = json.loads((Path(scratch) / "probe.json").read_text())
        print("Installed host:", info, flush=True)
        version = info["office"]
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(hive, fr"Software\Policies\Microsoft\Office\{version}\Excel\Security") as key:
                    if winreg.QueryValueEx(key, "AccessVBOM")[0] == 0:
                        raise RuntimeError("Managed policy disables AccessVBOM")
            except FileNotFoundError:
                pass
        changes = []
        try:
            settings = [(fr"Software\Microsoft\Office\{version}\Excel\Security", "AccessVBOM", 1)]
            # VBE version is read from the actual host once project access is available.
            for path, name, value in settings:
                old = read_value(path, name)
                changes.append((path, name, old))
                print("Original setting:", path, name, old, flush=True)
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
            if not info["access"]:
                bounded_worker("Excel", scratch, "probe", "")
                info = json.loads((Path(scratch) / "probe.json").read_text())
            if not info["access"] or not info["vbe"]:
                raise RuntimeError("Cannot determine actual VBE version")
            vbe_version = ".".join(str(int(part)) for part in info["vbe"].split("."))
            for name in ("BreakOnAllErrors", "BreakOnServerErrors"):
                path = fr"Software\Microsoft\VBA\{vbe_version}\Common"
                old = read_value(path, name)
                changes.append((path, name, old))
                print("Original error trapping:", path, name, old, flush=True)
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, 0)
            bounded_worker("Excel", scratch, "tests", "")
        finally:
            for path, name, old in reversed(changes):
                restore_value(path, name, old)
            print("Exact original registry values restored and verified.", flush=True)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        worker(*sys.argv[2:])
    else:
        main()
