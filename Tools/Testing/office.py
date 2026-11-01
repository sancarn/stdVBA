"""Shared bounded Office lifecycle with owned-process monitoring and exact registry restoration."""
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
            config = json.loads(os.environ["STDVBA_TEST_CONFIG"])
            app.DisplayAlerts = config.get("display_alerts", 0)
            project = doc.VBProject
            for relative in config["files"]:
                path = ROOT / relative
                raw = path.read_bytes()
                if b"\n" in raw.replace(b"\r\n", b"") or b"\r" in raw.replace(b"\r\n", b""):
                    raise RuntimeError("CRLF required before import: " + relative)
                project.VBComponents.Import(str(path))
            macro = "UnifiedTestRunner.RunCase"
            if host == "Excel":
                macro = "'" + doc.Name + "'!" + macro
            result = app.Run(macro)
            (Path(scratch) / "result.json").write_bytes(str(result).encode("utf-8"))
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
    environment = dict(os.environ)
    process = subprocess.Popen([sys.executable, "-B", __file__, "--worker", host, scratch, phase, url], env=environment)
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



def _run(config, host, scratch):
    os.environ["STDVBA_TEST_CONFIG"] = json.dumps(config)
    bounded_worker(host, scratch, "probe", "")
    info = json.loads((Path(scratch) / "probe.json").read_bytes())
    version = info["office"]
    access_path = r"Software\Microsoft\Office" + "\\" + version + "\\" + host + r"\Security"
    originals = {}
    try:
        if not info["access"]:
            # Respect either user or machine managed policy.
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                policy = r"Software\Policies\Microsoft\Office" + "\\" + version + "\\" + host + r"\Security"
                try:
                    with winreg.OpenKey(hive, policy) as key:
                        if winreg.QueryValueEx(key, "AccessVBOM")[0] == 0:
                            raise RuntimeError("Managed policy disables VBA project access for " + host)
                except FileNotFoundError:
                    pass
            originals[(access_path, "AccessVBOM")] = read_value(access_path, "AccessVBOM")
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, access_path) as key:
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            bounded_worker(host, scratch, "probe", "")
            info = json.loads((Path(scratch) / "probe.json").read_bytes())
            if not info["access"]:
                raise RuntimeError("VBA project access remains unavailable for " + host)
        trap_path = r"Software\Microsoft\VBA" + "\\" + ".".join(str(int(part)) for part in info["vbe"].split(".")) + r"\Common"
        for name in ("BreakOnAllErrors", "BreakOnServerErrors"):
            originals[(trap_path, name)] = read_value(trap_path, name)
        print(host + " original settings: " + repr(originals), flush=True)
        for name in ("BreakOnAllErrors", "BreakOnServerErrors"):
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, trap_path) as key:
                winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, 0)
        if config["kind"] == "adapter":
            command = ROOT / config["command"]
            args = [sys.executable, "-B", str(command)] if command.suffix == ".py" else ["powershell", "-NoProfile", "-File", str(command)]
            if config["id"] == "stdImage":
                args += ["--hosts", host]
            subprocess.run(args, cwd=ROOT, check=True, timeout=600)
        else:
            bounded_worker(host, scratch, "tests", "")
    finally:
        errors = []
        for (path, name), original in reversed(list(originals.items())):
            try:
                restore_value(path, name, original)
            except Exception as error:
                errors.append(str(error))
        if errors:
            raise RuntimeError("Settings restoration failed: " + "; ".join(errors))
        print(host + ": exact registry values restored", flush=True)


def run(config, host, scratch):
    # Registry settings are shared across Office instances. Refuse overlapping
    # runner transactions; Windows releases the byte lock if this process exits.
    import msvcrt
    lock_path = Path(tempfile.gettempdir()) / "stdVBA-office-tests.lock"
    with lock_path.open("a+b") as lock:
        lock.seek(0, 2)
        if lock.tell() == 0:
            lock.write(b"0")
            lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise RuntimeError("Another stdVBA Office test transaction is running") from error
        original = os.environ.get("STDVBA_TEST_CONFIG")
        try:
            _run(config, host, scratch)
        finally:
            if original is None:
                os.environ.pop("STDVBA_TEST_CONFIG", None)
            else:
                os.environ["STDVBA_TEST_CONFIG"] = original
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


if __name__ == "__main__":
    worker(*sys.argv[2:])
