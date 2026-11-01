"""Bounded isolated Office runner. Requires existing VBA project access.
Does not change Office security or debugging preferences.
"""
import ctypes
from ctypes import wintypes
import os
import sys
sys.dont_write_bytecode = True
from pathlib import Path
import socket
import subprocess
import tempfile
import threading
import time
from http.server import ThreadingHTTPServer
from fixture import Handler

root = Path(__file__).resolve().parents[3]
for name in ('src/stdICallable.cls', 'src/stdHTTP.cls', 'tests/stdHTTPLocalTests.bas'):
    raw = (root / name).read_bytes().replace(b'\r\n', b'')
    if b'\n' in raw or b'\r' in raw:
        raise SystemExit('Source must use CRLF before import: ' + name)

user = ctypes.WinDLL('user32', use_last_error=True)
callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]

with ThreadingHTTPServer(('127.0.0.1', 0), Handler) as server, socket.socket() as refused, tempfile.TemporaryDirectory() as scratch:
    # Reserve a non-listening endpoint so no unrelated service can take its port.
    refused.bind(('127.0.0.1', 0))
    threading.Thread(target=server.serve_forever, daemon=True).start()
    env = dict(os.environ, STDHTTP_URL='http://127.0.0.1:%d' % server.server_port,
               STDHTTP_REFUSED_URL='http://127.0.0.1:%d' % refused.getsockname()[1])
    pidfile = Path(scratch) / 'office.pid'
    process = subprocess.Popen(['powershell', '-NoProfile', '-File', str(Path(__file__).with_name('worker.ps1')), '-PidFile', str(pidfile)], env=env)
    owned_pid = None
    dialog = False
    deadline = time.monotonic() + 45
    try:
        while process.poll() is None:
            if pidfile.exists():
                value = pidfile.read_text().strip()
                if value: owned_pid = int(value)
            @callback_type
            def inspect(hwnd, unused):
                global dialog
                pid = wintypes.DWORD()
                user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if owned_pid and pid.value == owned_pid:
                    name = ctypes.create_unicode_buffer(256)
                    user.GetClassNameW(hwnd, name, 256)
                    if name.value == '#32770':
                        dialog = True
                        title = ctypes.create_unicode_buffer(1024)
                        user.GetWindowTextW(hwnd, title, 1024)
                        print('Office dialog (test failed):', title.value, flush=True)
                        user.PostMessageW(hwnd, 0x0010, 0, 0)
                    user.ShowWindow(hwnd, 6)
                return True
            user.EnumWindows(inspect, 0)
            if dialog or time.monotonic() >= deadline:
                raise RuntimeError('Office dialog detected' if dialog else 'Office test deadline exceeded')
            time.sleep(0.1)
        if process.returncode:
            raise RuntimeError('Office worker failed (%d)' % process.returncode)
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=5)
        if owned_pid:
            # Only the PID obtained from this worker's new Excel instance.
            subprocess.run(['taskkill', '/PID', str(owned_pid), '/F'], capture_output=True)
        server.shutdown()
