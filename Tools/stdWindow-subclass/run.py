"""Bounded isolated Excel regression runner; requires pywin32 and Office.

Discovers the installed Office/VBE versions, restores exact registry values,
and dismisses dialogs only in the newly created test process.
"""
import json
from pathlib import Path
import subprocess
import sys
import threading
import tempfile
import winreg
import gc
import os

ROOT = Path(__file__).resolve().parents[2]
FILES = ['src/stdICallable.cls', 'src/stdCallback.cls', 'src/stdWindow.cls',
         'tests/stdWindowSubclassProbe.cls', 'tests/stdWindowSubclassTests.bas',
         'Tools/stdWindow-subclass/IDEProbe.bas']


def worker():
    print('Starting isolated Office worker', flush=True)
    import win32com.client
    import win32process
    import win32gui
    import win32con
    existing = set()
    win32gui.EnumWindows(lambda hwnd, _: existing.add(win32process.GetWindowThreadProcessId(hwnd)[1]), None)
    app = win32com.client.DispatchEx('Excel.Application')
    print('Created isolated Office application', flush=True)
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    if pid in existing:
        raise RuntimeError('Worker returned an existing Office instance; refusing to use it.')
    Path(sys.argv[2]).write_text(str(pid), encoding='ascii')
    print('PID ' + str(pid), flush=True)
    stop = threading.Event()
    dialogs = []

    def monitor():
        while not stop.wait(.03):
            def visit(hwnd, unused):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                    return
                if win32gui.GetClassName(hwnd) == '#32770' and win32gui.IsWindowVisible(hwnd):
                    labels = []
                    win32gui.EnumChildWindows(hwnd, lambda child, _: labels.append(win32gui.GetWindowText(child)), None)
                    dialogs.append(' | '.join(labels))
                    print('DIALOG ' + dialogs[-1], flush=True)
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                elif win32gui.GetClassName(hwnd) in ('XLMAIN', 'wndclass_desked_gsk') and win32gui.IsWindowVisible(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            win32gui.EnumWindows(visit, None)

    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()
    book = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        no_vbom = '--no-vbom' in sys.argv
        if no_vbom:
            book = app.Workbooks.Open(sys.argv[3], ReadOnly=True)
            try:
                book.VBProject.VBComponents.Count
            except Exception as error:
                if 'not trusted' not in str(error).lower():
                    raise
                print('Verified VBOM access is disabled in this test instance', flush=True)
            else:
                raise RuntimeError('Negative access test did not disable VBOM')
        else:
            book = app.Workbooks.Add()
            project = book.VBProject
            project.References.AddFromGuid('{1EA4DBF0-3C3B-11CF-810C-00AA00389B71}', 1, 1)
            project.VBComponents.Add(3).Name = 'stdWindowTestForm'
            for file in FILES:
                raw = (ROOT / file).read_bytes().replace(b'\r\n', b'')
                assert b'\n' not in raw and b'\r' not in raw, file
                project.VBComponents.Import(str(ROOT / file))
            app.VBE.CommandBars.FindControl(1, 578).Execute()
            if dialogs:
                raise RuntimeError('Compilation failed')
            book.SaveAs(sys.argv[3], 52)
        result = app.Run("'" + book.Name + "'!stdWindowSubclassTests.RunTests")
        print(json.dumps({'result': result, 'dialogs': dialogs}), flush=True)
        if dialogs or not str(result).startswith('Passed '):
            raise RuntimeError('Office regression failed')
        result = app.Run("'" + book.Name + "'!stdWindowSubclassTests.RunUserFormTests")
        print(json.dumps({'result': result, 'dialogs': dialogs}), flush=True)
        if dialogs or not str(result).startswith('Passed '):
            raise RuntimeError('UserForm regression failed')
        result = app.Run("'" + book.Name + "'!stdWindowSubclassTests.RunResolverTests")
        print(json.dumps({'result': result, 'dialogs': dialogs}), flush=True)
        if dialogs or not str(result).startswith('Passed '):
            raise RuntimeError('EbMode resolver regression failed')
        if not no_vbom:
            ide_checks(app, book, project, pid)
        if dialogs:
            raise RuntimeError('Office/VBE dialog occurred during IDE checks')
    finally:
        try:
            if book is not None:
                print('Closing disposable workbook', flush=True)
                book.Close(False)
        finally:
            try:
                print('Closing isolated Office application', flush=True)
                app.Quit()
                print('Isolated Office application closed', flush=True)
            finally:
                stop.set()
                thread.join(2)
    # Office has closed and assertions have completed. In this disposable
    # process, skip late pywin32 proxy finalizers: after a denied VBOM query
    # they can block interpreter teardown even after Excel.Quit returned.
    # The parent still owns the Office PID and performs failure-safe cleanup.
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(0)


def ide_checks(app, book, project, pid):
    import ctypes
    from ctypes import wintypes
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user.SendMessageTimeoutW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
                                        wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t)]
    user.SendMessageTimeoutW.restype = ctypes.c_size_t
    user.DestroyWindow.argtypes = [wintypes.HWND]
    user.IsWindow.argtypes = [wintypes.HWND]
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.ReadProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p,
                                         ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    process = kernel.OpenProcess(0x10, False, pid)
    if not process:
        raise ctypes.WinError(ctypes.get_last_error())

    def run(name):
        return app.Run("'" + book.Name + "'!WindowIDEProbe." + name)

    def counts(thunk):
        counters = (wintypes.DWORD * 2)()
        copied = ctypes.c_size_t()
        if not kernel.ReadProcessMemory(process, thunk+4096, counters, ctypes.sizeof(counters), ctypes.byref(copied)):
            raise ctypes.WinError(ctypes.get_last_error())
        return tuple(counters)

    def send(hwnd):
        result = ctypes.c_size_t()
        if not user.SendMessageTimeoutW(hwnd, 0x8001, 1, 2, 2, 2000, ctypes.byref(result)):
            raise ctypes.WinError(ctypes.get_last_error())
        return result.value

    def paused_probe(hwnd, thunk, reset=False):
        import pythoncom
        import win32com.client
        import pywintypes
        import time
        import win32gui
        vbe_hwnd = app.VBE.MainWindow.HWnd
        run_id = reset_id = None
        for menu in app.VBE.CommandBars('Menu Bar').Controls:
            if menu.Caption.replace('&', '').lower() == 'run':
                for control in menu.Controls:
                    caption = control.Caption.replace('&', '').lower()
                    if caption.startswith('run sub/') or caption == 'run macro':
                        run_id = control.Id
                    if caption == 'reset':
                        reset_id = control.Id
        assert run_id and reset_id, 'Could not find the test VBE Continue/Reset controls'
        stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app.VBE._oleobj_)
        errors = []

        def coordinator():
            pythoncom.CoInitialize()
            remote = None
            try:
                remote = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch))
                deadline = time.monotonic()+8
                while time.monotonic() < deadline:
                    if '[break]' in win32gui.GetWindowText(vbe_hwnd).lower():
                        break
                    time.sleep(.03)
                else:
                    raise RuntimeError('Test VBE did not enter break mode: ' + win32gui.GetWindowText(vbe_hwnd))
                assert send(hwnd) == 0, 'Break mode must forward without entering VBA'
                assert counts(thunk) == (0, 1), 'Break mode should preserve the subscription'
                remote.CommandBars.FindControl(1, reset_id if reset else run_id).Execute()
            except Exception as error:
                errors.append(str(error))
                print('Break-mode coordinator failed:', error, flush=True)
                if remote is not None:
                    try:
                        remote.CommandBars.FindControl(1, reset_id).Execute()
                    except Exception:
                        pass
            finally:
                remote = None
                pythoncom.CoUninitialize()

        helper = threading.Thread(target=coordinator, daemon=True)
        helper.start()
        try:
            result = run('StopAndResume')
        except pywintypes.com_error as error:
            if not reset or not error.excepinfo or error.excepinfo[5] != -2146788248:
                raise
            result = None
        helper.join(10)
        assert not helper.is_alive(), 'Break-mode coordinator did not finish'
        assert not errors, errors
        if not reset:
            assert result == 2468, 'Continue must resume callback dispatch'

    try:
        hwnd, thunk = run('ArmModeless')
        assert send(hwnd) == 2468, 'Modeless callbacks must work after the starter macro returns'
        run('Cleanup')
        print('Passed real modeless UserForm callback after starter macro returns', flush=True)

        hwnd, thunk = run('Arm')
        print('IDE mode after macro:', project.Mode, 'native counters', counts(thunk), flush=True)
        assert send(hwnd) == 0, 'Design mode must forward without executing VBA'
        assert counts(thunk) == (0, 0), 'Design mode must remove the registration natively'
        assert run('CheckDetached'), 'Wrapper must reconcile native detachment'
        run('Cleanup')
        print('Passed real design-mode detachment', flush=True)

        hwnd, thunk = run('Arm')
        paused_probe(hwnd, thunk)
        run('Cleanup')
        print('Passed real break/Continue forwarding', flush=True)

        hwnd, thunk = run('Arm')
        paused_probe(hwnd, thunk, reset=True)
        if user.IsWindow(hwnd):
            assert send(hwnd) == 0
            assert counts(thunk) == (0, 0)
        print('Passed real VBE Reset detachment', flush=True)

        hwnd, thunk = run('Arm')
        import pywintypes
        try:
            run('EndNow')
        except pywintypes.com_error as error:
            # VBE returns its explicit execution-aborted HRESULT for End.
            if not error.excepinfo or error.excepinfo[5] != -2146788248:
                raise
        if user.IsWindow(hwnd):
            assert send(hwnd) == 0, 'End must not call a stale VBA object'
            assert counts(thunk) == (0, 0), 'End must detach without VBA globals'
            # The disposable Office process owns this orphan after End; process
            # teardown reclaims it. Do not enter its vanished VBA wrapper again.
        print('Passed real End/reset detachment', flush=True)
    finally:
        kernel.CloseHandle(process)


def launch_worker(template, no_vbom=False):
    handle = tempfile.NamedTemporaryFile(prefix='stdvba86-', suffix='.pid', delete=False)
    pidfile = Path(handle.name)
    handle.close()
    args = [sys.executable, __file__, '--worker', str(pidfile), str(template)]
    if no_vbom:
        args.append('--no-vbom')
    child = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    timed_out = False
    output = ''
    try:
        output, _ = child.communicate(timeout=45)
    except subprocess.TimeoutExpired:
        timed_out = True
        child.kill()
    finally:
        owned_pid = pidfile.read_text().strip()
        if owned_pid:
            subprocess.run(['taskkill', '/PID', owned_pid, '/F'], capture_output=True)
        if child.poll() is None:
            child.kill()
        tail, _ = child.communicate(timeout=5)
        if timed_out:
            output = tail
        pidfile.unlink()
    print(output, flush=True)
    if timed_out:
        raise RuntimeError('Office test timeout; partial worker output shown above')
    if child.returncode:
        raise RuntimeError('Office worker failed: ' + str(child.returncode))


def main():
    import win32com.client
    import win32process
    import win32gui
    existing = set()
    win32gui.EnumWindows(lambda hwnd, _: existing.add(win32process.GetWindowThreadProcessId(hwnd)[1]), None)
    app = win32com.client.DispatchEx('Excel.Application')
    discovery_pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    if discovery_pid in existing:
        raise RuntimeError('Discovery returned an existing Office instance; refusing to use it.')
    print('Discovery PID', discovery_pid, flush=True)
    try:
        version = app.Version
        vbe = '.'.join(str(int(part)) for part in app.VBE.Version.split('.'))
        book = app.Workbooks.Add()
        try:
            access = book.VBProject.VBComponents.Count > 0
        except Exception:
            access = False
        finally:
            book.Close(False)
    finally:
        app.Quit()
        del app
        gc.collect()
        subprocess.run(['taskkill', '/PID', str(discovery_pid), '/F'], capture_output=True)
    print('Excel', version, 'VBE', vbe, 'VBProject access', access, flush=True)
    policy_enables_access = False
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, fr'Software\Policies\Microsoft\Office\{version}\Excel\Security') as key:
                policy, _ = winreg.QueryValueEx(key, 'AccessVBOM')
                if not policy:
                    raise RuntimeError('Managed policy disables Excel VBA project access.')
                policy_enables_access = True
        except FileNotFoundError:
            pass
    template_handle = tempfile.NamedTemporaryFile(prefix='stdvba86-', suffix='.xlsm', delete=False)
    template = Path(template_handle.name)
    template_handle.close()
    changes = []
    try:
        settings = [(fr'Software\Microsoft\VBA\{vbe}\Common', 'BreakOnAllErrors', 0),
                    (fr'Software\Microsoft\VBA\{vbe}\Common', 'BreakOnServerErrors', 0)]
        if not access:
            settings.append((fr'Software\Microsoft\Office\{version}\Excel\Security', 'AccessVBOM', 1))
        for path, name, value in settings:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
                try:
                    old, kind = winreg.QueryValueEx(key, name)
                except FileNotFoundError:
                    old, kind = None, None
                changes.append((path, name, old, kind))
                print('SETTING', path, name, old, kind, flush=True)
                winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
        launch_worker(template)
        if policy_enables_access:
            print('Skipped disabled-VBOM variant: managed policy forces access on.', flush=True)
        else:
            path = fr'Software\Microsoft\Office\{version}\Excel\Security'
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
                try:
                    old, kind = winreg.QueryValueEx(key, 'AccessVBOM')
                except FileNotFoundError:
                    old, kind = None, None
                changes.append((path, 'AccessVBOM', old, kind))
                print('SETTING disabled-VBOM variant', path, 'AccessVBOM', old, kind, flush=True)
                winreg.SetValueEx(key, 'AccessVBOM', 0, winreg.REG_DWORD, 0)
            launch_worker(template, no_vbom=True)
    finally:
        for path, name, old, kind in reversed(changes):
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
                if kind is None:
                    winreg.DeleteValue(key, name)
                else:
                    winreg.SetValueEx(key, name, 0, kind, old)
        template.unlink(missing_ok=True)
        print('Original registry values restored.', flush=True)


if __name__ == '__main__':
    worker() if '--worker' in sys.argv else main()
