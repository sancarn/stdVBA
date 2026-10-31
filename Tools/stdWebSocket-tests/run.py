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
import asyncio
import ssl
import shutil
from contextlib import contextmanager
from fixture import start

root = Path(__file__).resolve().parents[2]
for name in ('src/stdWebSocket.cls', 'src/WIP/stdWebSocket/stdWebSocketNew.cls', 'src/stdICallable.cls', 'Tools/stdWebSocket-tests/WebSocketRecorder.cls', 'tests/stdWebSocketLocalTests.bas', 'Tools/stdWebSocket-tests/WebSocketAsyncRecorder.cls', 'tests/stdWebSocketAsyncTests.bas'):
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
user.EnumChildWindows.argtypes = [wintypes.HWND, callback_type, wintypes.LPARAM]
user.IsWindowVisible.argtypes = [wintypes.HWND]
user.IsWindowVisible.restype = wintypes.BOOL

@contextmanager
def certificate():
    """Untrusted TLS fixture. Never reads or modifies certificate trust stores."""
    openssl = shutil.which('openssl') or r'C:\Program Files\Git\usr\bin\openssl.exe'
    with tempfile.TemporaryDirectory() as folder:
        folder = Path(folder)
        cert, key, der = (folder / n for n in ('cert.pem', 'key.pem', 'cert.cer'))
        config = folder / 'openssl.cnf'
        config.write_bytes(b'[req]\r\ndistinguished_name=dn\r\nx509_extensions=ext\r\n[dn]\r\n[ext]\r\nsubjectAltName=DNS:localhost\r\nbasicConstraints=critical,CA:FALSE\r\nkeyUsage=critical,digitalSignature,keyEncipherment\r\nextendedKeyUsage=serverAuth\r\n')
        subprocess.run([openssl, 'req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', str(key), '-out', str(cert), '-days', '1', '-subj', '/CN=localhost', '-config', str(config)], check=True, capture_output=True)
        subprocess.run([openssl, 'x509', '-in', str(cert), '-outform', 'DER', '-out', str(der)], check=True, capture_output=True)
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(cert, key)
        yield context

native = Path(os.environ['TEMP']) / 'stdWebSocket-native-build'
subprocess.run(['powershell', '-NoProfile', '-File', str(root / 'src/WIP/stdWebSocket/native/build.ps1'), '-OutputDirectory', str(native), '-WithTests'], check=True, timeout=60)

with socket.socket() as refused, tempfile.TemporaryDirectory() as scratch, certificate() as tls:
    # Reserve a non-listening endpoint so no unrelated service can take its port.
    refused.bind(('127.0.0.1', 0))
    ready = threading.Event()
    ports = []
    def report(port):
        ports.append(port)
        ready.set()
    threading.Thread(target=lambda: asyncio.run(start(report)), daemon=True).start()
    if not ready.wait(5): raise RuntimeError("Fixture startup timed out")
    ready.clear()
    threading.Thread(target=lambda: asyncio.run(start(report, tls)), daemon=True).start()
    if not ready.wait(5): raise RuntimeError('TLS fixture startup timed out')
    for architecture in ('x86', 'x64'):
        subprocess.run([str(native / architecture / 'native_tests.exe'), str(ports[0])], check=True, timeout=30)
    env = dict(os.environ, STDWEBSOCKET_NATIVE_ROOT=str(native), STDWEBSOCKET_TLS='wss://localhost:%d' % ports[1], STDWEBSOCKET_URL='ws://127.0.0.1:%d' % ports[0],
               STDWEBSOCKET_REFUSED_URL='ws://127.0.0.1:%d' % refused.getsockname()[1])
    pidfile = Path(scratch) / 'office.pid'
    env['STDWEBSOCKET_TRACE'] = str(Path(scratch) / 'progress.txt')
    process = subprocess.Popen(['powershell', '-NoProfile', '-File', str(Path(__file__).with_name('worker.ps1')), '-PidFile', str(pidfile)], env=env)
    owned_pid = None
    dialog = False
    deadline = time.monotonic() + 45
    last_trace = ''
    try:
        while process.poll() is None:
            tracefile = Path(env['STDWEBSOCKET_TRACE'])
            if tracefile.exists():
                trace = tracefile.read_text(errors='replace').strip()
                if trace and trace != last_trace:
                    print('VBA:', trace, flush=True)
                    last_trace = trace
            if pidfile.exists():
                value = pidfile.read_text().strip()
                if value: owned_pid = int(value)
            @callback_type
            def inspect(hwnd, unused):
                global dialog
                pid = wintypes.DWORD()
                user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if owned_pid and pid.value == owned_pid and user.IsWindowVisible(hwnd):
                    name = ctypes.create_unicode_buffer(256)
                    user.GetClassNameW(hwnd, name, 256)
                    if name.value == '#32770':
                        dialog = True
                        title = ctypes.create_unicode_buffer(1024)
                        user.GetWindowTextW(hwnd, title, 1024)
                        captions = []
                        @callback_type
                        def child_text(child, unused):
                            caption = ctypes.create_unicode_buffer(2048)
                            user.GetWindowTextW(child, caption, 2048)
                            if caption.value: captions.append(caption.value)
                            return True
                        user.EnumChildWindows(hwnd, child_text, 0)
                        print('Office dialog (test failed):', title.value, captions, flush=True)
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

