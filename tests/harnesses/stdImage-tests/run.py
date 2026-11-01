"""Run bounded stdImage regressions in disposable Excel and Word instances.

Requires Windows, Office, pywin32 and Pillow. The parent monitors only worker-owned
Office windows, restores exact registry values and restores the clipboard.
"""
import argparse
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import threading
import time
import winreg
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[3]
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
            for relative in ("src/stdImage.cls", "tests/stdImageTests.bas"):
                raw = (ROOT / relative).read_bytes()
                if b"\n" in raw.replace(b"\r\n", b"") or b"\r" in raw.replace(b"\r\n", b""):
                    raise RuntimeError("CRLF required before import: " + relative)
                print("Importing " + relative, flush=True)
                project.VBComponents.Import(str(ROOT / relative))
                print("Imported " + relative, flush=True)
            # Running invokes the compiler; modal failures are caught by the parent.
            macro = "'" + doc.Name + "'!stdImageTests.RunTests" if host == "Excel" else "stdImageTests.RunTests"
            print("Running " + macro, flush=True)
            result = app.Run(macro, scratch, url)
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
        if phase == "tests":
            leftovers = [p.name for p in exports.iterdir() if p.suffix.lower() in (".png", ".bmp", ".gif", ".jpeg", ".tiff")]
            if leftovers:
                raise RuntimeError("Image temporary files leaked: " + repr(leftovers))
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=5)
        if owned:
            # Only the PID reported by our DispatchEx-created instance.
            subprocess.run(["taskkill", "/PID", str(owned), "/F"], capture_output=True)


def fixtures(folder):
    from PIL import Image
    image = Image.new("RGB", (3, 2))
    image.putdata([(255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 255), (0, 0, 0), (255, 255, 0)])
    for ext, fmt in (("png", "PNG"), ("bmp", "BMP"), ("gif", "GIF"), ("jpg", "JPEG"), ("tif", "TIFF")):
        image.save(folder / ("fixture." + ext), fmt)
    transparent = Image.new("RGBA", (2, 2), (0, 0, 0, 0))
    transparent.putpixel((1, 1), (255, 0, 0, 255))
    transparent.save(folder / "transparent.png")
    large = Image.new("RGB", (64, 64))
    large.putdata([((x * 13 + y * 7) % 256, (x * 19) % 256, (y * 23) % 256) for y in range(64) for x in range(64)])
    large.save(folder / "quality.png")
    image.resize((16, 16)).save(folder / "fixture.ico", bitmap_format="bmp")
    # Minimal standard WMF with a red filled rectangle in a 30 x 20 window.
    def record(function, *words):
        return struct.pack("<IH" + "H" * len(words), 3 + len(words), function, *words)
    records = [
        record(0x0103, 8), record(0x020B, 0, 0), record(0x020C, 20, 30),
        record(0x02FC, 0, 255, 0, 0), record(0x012D, 0),
        record(0x041B, 20, 30, 0, 0), record(0x0000),
    ]
    data = b"".join(records)
    header = struct.pack("<HHHIHIH", 1, 9, 0x0300, 9 + len(data) // 2, 1, 7, 0)
    # Placeable header supplies physical bounds to GDI+.
    placeable = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, 0, 0, 30, 20, 1440, 0)
    checksum = 0
    for word in struct.unpack("<10H", placeable):
        checksum ^= word
    (folder / "fixture.wmf").write_bytes(placeable + struct.pack("<H", checksum) + header + data)
    gdi = ctypes.WinDLL("gdi32")
    gdi.CreateEnhMetaFileW.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.POINTER(wintypes.RECT), wintypes.LPCWSTR]
    gdi.CreateEnhMetaFileW.restype = wintypes.HDC
    gdi.CloseEnhMetaFile.argtypes = [wintypes.HDC]
    gdi.CloseEnhMetaFile.restype = wintypes.HANDLE
    gdi.Rectangle.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
    gdi.DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
    bounds = wintypes.RECT(0, 0, 300, 200)
    dc = gdi.CreateEnhMetaFileW(None, str(folder / "fixture.emf"), ctypes.byref(bounds), None)
    if not dc:
        raise ctypes.WinError()
    gdi.Rectangle(dc, 0, 0, 30, 20)
    handle = gdi.CloseEnhMetaFile(dc)
    if not handle:
        raise ctypes.WinError()
    gdi.DeleteEnhMetaFile(handle)


def main(hosts):
    import pythoncom
    ole = ctypes.OleDLL("ole32")
    result = ole.OleInitialize(None)
    if result < 0:
        raise RuntimeError("Could not initialise OLE clipboard access")
    clipboard = pythoncom.OleGetClipboard()
    try:
        with tempfile.TemporaryDirectory(prefix="stdImage-98-") as temp:
            folder = Path(temp)
            fixtures(folder)

            class Handler(BaseHTTPRequestHandler):
                def do_GET(self):
                    if self.path == "/image":
                        data = (folder / "fixture.png").read_bytes()
                        self.send_response(200)
                        self.send_header("Content-Type", "image/png")
                        self.send_header("Content-Length", str(len(data)))
                        self.end_headers()
                        self.wfile.write(data)
                    else:
                        self.send_error(404)
                def log_message(self, *args):
                    pass

            with ThreadingHTTPServer(("127.0.0.1", 0), Handler) as server:
                threading.Thread(target=server.serve_forever, daemon=True).start()
                url = "http://127.0.0.1:" + str(server.server_port)
                try:
                    for host in hosts:
                        bounded_worker(host, temp, "probe", url)
                        info = json.loads((folder / "probe.json").read_bytes())
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
                                bounded_worker(host, temp, "probe", url)
                                info = json.loads((folder / "probe.json").read_bytes())
                                if not info["access"]:
                                    raise RuntimeError("VBA project access remains unavailable for " + host)
                            trap_path = r"Software\Microsoft\VBA" + "\\" + ".".join(str(int(part)) for part in info["vbe"].split(".")) + r"\Common"
                            for name in ("BreakOnAllErrors", "BreakOnServerErrors"):
                                originals[(trap_path, name)] = read_value(trap_path, name)
                            print(host + " original settings: " + repr(originals), flush=True)
                            for name in ("BreakOnAllErrors", "BreakOnServerErrors"):
                                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, trap_path) as key:
                                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, 0)
                            bounded_worker(host, temp, "tests", url)
                        finally:
                            for (path, name), original in reversed(list(originals.items())):
                                restore_value(path, name, original)
                            print(host + ": exact registry values restored", flush=True)
                finally:
                    server.shutdown()
    finally:
        for attempt in range(20):
            try:
                pythoncom.OleSetClipboard(clipboard)
                pythoncom.OleFlushClipboard()
                break
            except pythoncom.com_error:
                if attempt == 19:
                    raise
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
        clipboard = None
        ole.OleUninitialize()
        print("Clipboard restored", flush=True)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        worker(*sys.argv[2:])
    else:
        parser = argparse.ArgumentParser()
        parser.add_argument("--hosts", nargs="+", default=["Excel", "Word"], choices=["Excel", "Word"])
        main(parser.parse_args().hosts)
