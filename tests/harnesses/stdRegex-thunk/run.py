"""Disposable Word x64 test host with bounded watchdog and exact registry cleanup.

Requires Python/pywin32 and desktop Word. Run from any directory. No user Office
instance is attached to or closed. The supervisor owns settings and timeout.
"""
from pathlib import Path
import ctypes
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import winreg

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]

def write(path, text):
    path.write_bytes(text.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8'))

def worker(folder, test_only=False):
    import win32com.client
    import win32process
    app = doc = None
    result = {'measured_at_utc': __import__('datetime').datetime.now(__import__('datetime').timezone.utc).isoformat()}
    try:
        app = win32com.client.DispatchEx('Word.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Add()
        pid = win32process.GetWindowThreadProcessId(doc.ActiveWindow.Hwnd)[1]
        (folder/'pid').write_text(str(pid))
        result['office_version'] = app.Version
        result['office_build'] = app.Build
        result['cpu'] = __import__('platform').processor()
        result['windows'] = __import__('platform').platform()
        project = doc.VBProject  # Check actual application's project access.
        paths = [folder/'stdRegex.cls',
                 ROOT/'tests/harnesses/stdRegex-build/TestReporter.cls',
                 ROOT/'tests/harnesses/stdRegex-build/TestRunner.bas', folder/'stdRegexTests.bas',
                 folder/'ThunkTests.bas', folder/'ThunkTestRunner.bas', HERE/'Benchmark.bas']
        for path in paths:
            data = path.read_bytes()
            assert b'\n' not in data.replace(b'\r\n', b'') and b'\r' not in data.replace(b'\r\n', b''), f'Not CRLF: {path}'
            project.VBComponents.Import(str(path))
        for name in ['ThunkBenchmark.Differential','StdRegexTestRunner.RunTests','ThunkTestRunner.RunTests']:
            print('Running '+name, flush=True)
            value = app.Run(name)
            result[name] = value
            if isinstance(value, str) or value <= 0:
                raise RuntimeError(f'{name}: {value}')
        result['samples'] = []
        for case in range(0 if test_only else 6):
            print('Benchmark case '+str(case), flush=True)
            rows = app.Run('ThunkBenchmark.RunCase', case, 7, 0.05)
            if isinstance(rows, str): raise RuntimeError(rows)
            for row in rows:
                result['samples'].append(dict(zip(['case','matches','captures','characters','operation','sample','iterations','elapsed_ms','ms_per_call'],row)))
        result['status'] = 'passed'
    except Exception as exc:
        result['status'] = 'failed'
        result['error'] = repr(exc)
        raise
    finally:
        write(folder/'result.json', json.dumps(result,indent=2)+'\n')
        try:
            if doc is not None: doc.Close(0)
        finally:
            if app is not None: app.Quit(0)

def supervise(test_only=False):
    import win32gui
    import win32process
    import win32con
    import psutil
    # Refuse managed AccessVBOM settings rather than overriding policy.
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CurVer') as key:
        version = winreg.QueryValueEx(key, '')[0].rsplit('.',1)[1]+'.0'
    for hive in (winreg.HKEY_CURRENT_USER,winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive,rf'Software\Policies\Microsoft\Office\{version}\Word\Security') as key:
                value,_ = winreg.QueryValueEx(key,'AccessVBOM')
                if not value: raise RuntimeError('Managed policy disables VBA project access')
        except FileNotFoundError: pass
    settings = [(rf'Software\Microsoft\Office\{version}\Word\Security','AccessVBOM',1),
                (r'Software\Microsoft\VBA\7.1\Common','BreakOnAllErrors',0),
                (r'Software\Microsoft\VBA\7.1\Common','BreakOnServerErrors',0)]
    saved=[]
    output=ROOT/'Tools/stdRegex-thunk'/('verification-x64' if test_only else 'results')
    output.mkdir(exist_ok=True)
    started=time.time()
    original_pids={p.pid for p in psutil.process_iter()}
    dialogs=[]
    process=None
    office_pid=None
    with tempfile.TemporaryDirectory(prefix='stdRegex-thunk-test-') as tmp:
        folder=Path(tmp)
        combined=(ROOT/'src/stdRegex.cls').read_text(encoding='utf-8')
        combined += '\nPublic Function CreateNative(ByVal pattern As String, Optional ByVal flags As String = "", Optional ByVal stepsLimit As Long = 10000) As stdRegex\n  Set CreateNative = Create(pattern, flags, stepsLimit, True)\nEnd Function\n'
        combined += '\nPublic Function CreateVBA(ByVal pattern As String, Optional ByVal flags As String = "", Optional ByVal stepsLimit As Long = 10000) As stdRegex\n  Set CreateVBA = Create(pattern, flags, stepsLimit, False)\nEnd Function\n'
        write(folder/'stdRegex.cls', combined)
        baseline_tests=(ROOT/'tests/stdRegexTests.bas').read_text(encoding='utf-8')
        write(folder/'stdRegexTests.bas', baseline_tests.replace('stdRegex.Create(', 'stdRegex.CreateVBA('))
        source=(ROOT/'tests/stdRegexTests.bas').read_text(encoding='utf-8')
        source=source.replace('stdRegex.Create(', 'stdRegex.CreateNative(')
        source=re.sub(r'\bstdRegexTests\b','ThunkTests',source)
        write(folder/'ThunkTests.bas',source)
        source=(ROOT/'tests/harnesses/stdRegex-build/TestRunner.bas').read_text(encoding='utf-8')
        write(folder/'ThunkTestRunner.bas',source.replace('StdRegexTestRunner','ThunkTestRunner').replace('stdRegexTests','ThunkTests'))
        try:
            for path,name,value in settings:
                try:
                    key=winreg.OpenKey(winreg.HKEY_CURRENT_USER,path,0,winreg.KEY_READ|winreg.KEY_WRITE)
                    existed=True
                except FileNotFoundError:
                    key=winreg.CreateKey(winreg.HKEY_CURRENT_USER,path); existed=False
                with key:
                    try: old,kind=winreg.QueryValueEx(key,name); present=True
                    except FileNotFoundError: old=kind=None; present=False
                    saved.append(dict(path=path,name=name,key_existed=existed,present=present,value=old,kind=kind))
                    winreg.SetValueEx(key,name,0,winreg.REG_DWORD,value)
            write(output/'settings-before.json',json.dumps(saved,indent=2)+'\n')
            process=subprocess.Popen([sys.executable,str(Path(__file__)), '--worker',str(folder)]+(['--test-only'] if test_only else []))
            deadline=time.monotonic()+240
            while process.poll() is None:
                if (folder/'pid').exists(): office_pid=int((folder/'pid').read_text())
                if office_pid and office_pid not in original_pids:
                    def visit(hwnd,_):
                        if win32process.GetWindowThreadProcessId(hwnd)[1]!=office_pid: return
                        if not win32gui.IsWindowVisible(hwnd): return
                        klass=win32gui.GetClassName(hwnd)
                        if klass=='#32770':
                            texts=[]; buttons=[]
                            def child(ch,_):
                                texts.append(win32gui.GetWindowText(ch))
                                if win32gui.GetClassName(ch)=='Button': buttons.append(ch)
                            win32gui.EnumChildWindows(hwnd,child,None)
                            dialogs.append(' | '.join(texts))
                            if buttons: win32gui.PostMessage(buttons[0],win32con.BM_CLICK,0,0)
                            else: win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
                        ctypes.windll.user32.ShowWindowAsync(hwnd,6)
                    win32gui.EnumWindows(visit,None)
                if time.monotonic()>deadline:
                    raise TimeoutError('Office test exceeded 240 seconds')
                time.sleep(0.05)
            if (folder/'result.json').exists():
                data=json.loads((folder/'result.json').read_text())
            else: data={'status':'failed','error':'No worker result'}
            data.update(dialogs=dialogs,elapsed_seconds=time.time()-started,
                        baseline_commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),
                        native_class_sha256=hashlib.sha256((ROOT/'src/stdRegex.cls').read_bytes()).hexdigest(),
                        architecture='Windows x64',methodology='QPC; 7 rotated samples; warmed cached patterns; calibrated >=50ms batches; allocation/release included; output equality checked')
            if dialogs: data['status']='failed'
            write(output/'results.json',json.dumps(data,indent=2)+'\n')
            if process.returncode or data['status']!='passed': raise RuntimeError(data.get('error',str(dialogs)))
            print('All requested checks passed',flush=True)
        finally:
            cleanup_error = None
            try:
                if process and process.poll() is None: process.kill(); process.wait(timeout=10)
                # Close only the newly created test process, including after a COM hang.
                if office_pid and office_pid not in original_pids and psutil.pid_exists(office_pid):
                    p=psutil.Process(office_pid)
                    if p.name().lower()=='winword.exe': p.terminate(); p.wait(timeout=10)
            except Exception as exc:
                cleanup_error = repr(exc)
            for entry in reversed(saved):
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,entry['path'],0,winreg.KEY_WRITE) as key:
                    if entry['present']: winreg.SetValueEx(key,entry['name'],0,entry['kind'],entry['value'])
                    else:
                        try: winreg.DeleteValue(key,entry['name'])
                        except FileNotFoundError: pass
                if not entry['key_existed']:
                    try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER,entry['path'])
                    except OSError: pass
            verified=[]
            for entry in saved:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER,entry['path']) as key:
                        value,kind=winreg.QueryValueEx(key,entry['name'])
                    assert entry['present'] and value==entry['value'] and kind==entry['kind']
                except FileNotFoundError:
                    assert not entry['present']
                verified.append(entry['name'])
            write(output/'settings-restored.json',json.dumps({'verified':verified,'cleanup_error':cleanup_error},indent=2)+'\n')
            if cleanup_error: raise RuntimeError('Office cleanup failed: '+cleanup_error)

if __name__=='__main__':
    if len(sys.argv)>1 and sys.argv[1]=='--worker': worker(Path(sys.argv[2]), '--test-only' in sys.argv)
    else: supervise('--test-only' in sys.argv)
