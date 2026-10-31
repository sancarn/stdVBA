"""Run isolated issue 85 probes; requires Windows, Excel and pywin32."""
import json
from pathlib import Path
import subprocess
import sys
import threading
import winreg

ROOT = Path(__file__).resolve().parents[2]
FILES = ['src/stdICallable.cls', 'src/stdCallback.cls', 'Tools/stdCallback-tests/Probe.bas']
NAMES = ['Object ByVal inferred', 'Object ByVal explicit', 'Object ByRef inferred',
         'Object ByRef explicit', 'Object ByRef VarPtr', 'Variant ByVal explicit',
         'Collection ByVal inferred', 'Object ByVal RunEx', 'Nothing ByVal', 'Dictionary ByVal']

def worker(case):
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    print('PID ' + str(pid), flush=True)
    stop = threading.Event()
    dialogs = []
    def monitor():
        import win32gui
        import win32con
        while not stop.wait(.05):
            def visit(hwnd, unused):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                    return
                if win32gui.GetClassName(hwnd) == '#32770' and win32gui.IsWindowVisible(hwnd):
                    labels = []
                    win32gui.EnumChildWindows(hwnd, lambda child, _: labels.append(win32gui.GetWindowText(child)), None)
                    dialogs.append(' | '.join(labels))
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                elif win32gui.IsWindowVisible(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            win32gui.EnumWindows(visit, None)
    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()
    book = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        book = app.Workbooks.Add()
        for file in FILES:
            book.VBProject.VBComponents.Import(str(ROOT / file))
        result = app.Run("'" + book.Name + "'!CallbackProbe.Probe", case)
        print(json.dumps({'result': result, 'dialogs': dialogs}), flush=True)
    finally:
        try:
            if book is not None:
                book.Close(False)
        finally:
            try:
                app.Quit()
            finally:
                stop.set()
                thread.join(2)

def main():
    for file in FILES:
        data = (ROOT / file).read_bytes()
        assert b'\n' not in data.replace(b'\r\n', b'') and b'\r' not in data.replace(b'\r\n', b''), file
    # Read the installed Excel version from its registered COM application.
    import win32com.client
    app = win32com.client.DispatchEx('Excel.Application')
    version = app.Version
    app.Quit()
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, fr'Software\Policies\Microsoft\Office\{version}\Excel\Security') as key:
                policy, _ = winreg.QueryValueEx(key, 'AccessVBOM')
                if not policy:
                    raise RuntimeError('Managed policy disables Excel VBA project access.')
        except FileNotFoundError:
            pass
    changes = []
    try:
        for path, name, value in [
            (fr'Software\Microsoft\Office\{version}\Excel\Security', 'AccessVBOM', 1),
            (r'Software\Microsoft\VBA\7.1\Common', 'BreakOnAllErrors', 0),
            (r'Software\Microsoft\VBA\7.1\Common', 'BreakOnServerErrors', 0)]:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
                try:
                    old, kind = winreg.QueryValueEx(key, name)
                except FileNotFoundError:
                    old, kind = None, None
                changes.append((path, name, old, kind))
                print('SETTING', path, name, old, kind, flush=True)
                winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
        for case, name in enumerate(NAMES):
            child = subprocess.Popen([sys.executable, __file__, str(case)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            first = child.stdout.readline().strip()
            try:
                output, _ = child.communicate(timeout=25)
            except subprocess.TimeoutExpired:
                child.kill()
                if first.startswith('PID '):
                    subprocess.run(['taskkill', '/PID', first[4:], '/F'], capture_output=True)
                output, _ = child.communicate()
                output += ' TIMEOUT'
            # A bad native call can disconnect COM before Quit succeeds.
            # The PID was obtained from this worker's DispatchEx instance only.
            if first.startswith('PID '):
                subprocess.run(['taskkill', '/PID', first[4:], '/F'], capture_output=True)
            print(name, first, output.strip(), 'exit', child.returncode, flush=True)
    finally:
        for path, name, old, kind in reversed(changes):
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
                if kind is None:
                    winreg.DeleteValue(key, name)
                else:
                    winreg.SetValueEx(key, name, 0, kind, old)
        print('Original registry values restored.', flush=True)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        worker(int(sys.argv[1]))
    else:
        main()
