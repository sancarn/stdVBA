"""Run registered stdVBA suites in disposable Office projects."""
import argparse
import json
from pathlib import Path
import sys
import tempfile
import time

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]


def load_manifest():
    data = json.loads((ROOT / 'tests/manifest.json').read_bytes())
    seen = set()
    for suite in data['suites']:
        if suite['id'].casefold() in seen:
            raise ValueError('Duplicate suite: ' + suite['id'])
        seen.add(suite['id'].casefold())
        for file in suite.get('files', []) + ([suite['command']] if 'command' in suite else []):
            path = (ROOT / file).resolve()
            if not path.is_relative_to(ROOT) or not path.is_file():
                raise ValueError('Invalid manifest file: ' + file)
        if not suite['cases']:
            raise ValueError('Empty suite: ' + suite['id'])
    return data['suites']


def select(suites, names=None, case=None, all_suites=False):
    wanted = {name.casefold() for name in names or []}
    unknown = wanted - {s['id'].casefold() for s in suites}
    if unknown:
        raise ValueError('Unknown suite: ' + ', '.join(sorted(unknown)))
    if case and not wanted:
        raise ValueError('--test requires --suite')
    selected = []
    for suite in suites:
        if not (suite['id'].casefold() in wanted if wanted else all_suites or suite['default']):
            continue
        cases = [name for name in suite['cases'] if not case or name.casefold() == case.casefold()]
        if case and not cases:
            raise ValueError('Unknown test for ' + suite['id'] + ': ' + case)
        selected.extend((suite, name) for name in cases)
    if not selected:
        raise ValueError('No tests selected')
    return selected


def bridge(suite, case):
    target = suite['cases'][case]
    mode = suite.get('result', 'assertions')
    call = '  Test.FullTest = ' + ('True' if suite.get('full_test') else 'False') + '\n  Call ' + target + '\n'
    if mode != 'assertions':
        call = '  Dim result As Variant\n  result = ' + target + '\n'
        condition = {'pass': 'Left$(CStr(result), 4) = "PASS"',
                     'count': 'IsNumeric(result) And result > 0',
                     'timer': 'result = True' if case == 'testInvalidResolution' else 'Left$(CStr(result), 4) = "PASS"'}[mode]
        call += '  Test.Assert "Returned result: " & CStr(result), ' + condition + '\n'
    return '''Attribute VB_Name = "UnifiedTestRunner"
Option Explicit
Public Function RunCase() As String
  On Error GoTo Failed
  Test.Refresh
'''+call+'''  If Test.checks = 0 Then Err.Raise 5, , "No assertions ran"
  RunCase = "{""passed"":" & LCase$(CStr(Test.failures = 0)) & ",""assertions"":" & CStr(Test.checks) & ",""failures"":" & CStr(Test.failures) & ",""detail"":" & Quote(Test.lastCheck) & "}"
  Exit Function
Failed:
  RunCase = "{""passed"":false,""detail"":" & Quote("VBA error " & Err.Number & ": " & Err.Description) & "}"
End Function
Private Function Quote(ByVal value As String) As String
  value = Replace(value, "\\", "\\\\")
  value = Replace(value, Chr$(34), "\\" & Chr$(34))
  value = Replace(value, vbCr, "\\r")
  value = Replace(value, vbLf, "\\n")
  value = Replace(value, vbTab, "\\t")
  Quote = Chr$(34) & value & Chr$(34)
End Function
'''


def execute(suite, case, host, clang=None, full=False):
    if suite.get('unsupported'):
        return {'status': 'skipped', 'detail': suite['unsupported']}
    if host not in suite['hosts']:
        return {'status': 'skipped', 'detail': 'Supported hosts: ' + ', '.join(suite['hosts'])}
    if suite['kind'] == 'command':
        import shutil
        import subprocess
        compiler = clang or shutil.which('clang')
        if suite.get('compiler') and not compiler:
            return {'status': 'skipped', 'detail': 'Requires llvm-mingw clang; pass --clang PATH'}
        args = [sys.executable, '-B', str(ROOT / suite['command'])]
        if suite.get('compiler'):
            args += ['--clang', compiler]
        subprocess.run(args, cwd=ROOT, check=True, timeout=300)
        return {'status': 'passed'}
    if suite['kind'] == 'python':
        import subprocess
        subprocess.run([sys.executable, '-B', str(ROOT / suite['command']), suite['cases'][case]], cwd=ROOT, check=True, timeout=300)
        return {'status': 'passed'}
    import office
    with tempfile.TemporaryDirectory(prefix='stdVBA-tests-') as scratch:
        config = dict(suite, full_test=full)
        if suite['kind'] == 'vba':
            path = Path(scratch) / 'UnifiedTestRunner.bas'
            path.write_bytes(bridge(config, case).replace('\n', '\r\n').encode('utf-8'))
            config['files'] = ['tests/support/Test.cls'] + suite['files'] + [str(path)]
            if suite.get('instrument') == 'date':
                source = (ROOT / 'src/WIP/stdDate.cls').read_text()
                source = source.replace('  Set CName = CachedNames(key)', '  Set CName = CachedNames(key)\n  StdRegexTestRunner.regexCount = CachedRegex.Count\n  StdRegexTestRunner.namesCount = CachedNames.Count')
                instrumented = Path(scratch) / 'stdDate.cls'
                instrumented.write_bytes(source.replace('\n', '\r\n').encode('utf-8'))
                config['files'][1] = str(instrumented)
            if suite.get('instrument') == 'lambda-vm':
                source = (ROOT / 'src/stdLambda.cls').read_text(encoding='utf-8')
                fragment = (ROOT / 'tests/support/stdLambdaVM.cls.inc').read_text(encoding='utf-8')
                lines = []
                for number, line in enumerate(fragment.splitlines(), 1):
                    if line.lstrip().startswith('Debug.Assert '):
                        expression = line.strip()[len('Debug.Assert '):]
                        line = '    Test.Assert "VM instruction line ' + str(number) + '", ' + expression
                    lines.append(line)
                instrumented = Path(scratch) / 'stdLambda.cls'
                instrumented.write_bytes((source + '\n' + '\n'.join(lines)).replace('\n', '\r\n').encode('utf-8'))
                config['files'][config['files'].index('src/stdLambda.cls')] = str(instrumented)
        try:
            office.run(config, host, scratch)
        except Exception as error:
            result = Path(scratch) / 'result.json'
            if result.exists():
                detail = json.loads(result.read_bytes())
                detail['passed'] = False
                detail['detail'] = detail.get('detail', '') + '; ' + str(error)
                return dict(detail, status='failed')
            raise error
        result = Path(scratch) / 'result.json'
        if result.exists():
            detail = json.loads(result.read_bytes())
            return dict(detail, status='passed' if detail['passed'] else 'failed')
        if suite['kind'] == 'vba':
            raise RuntimeError('Office worker did not produce a result')
        return {'status': 'passed'}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', action='append', help='Suite ID; repeat to select several')
    parser.add_argument('--test', help='Exact case ID from --list')
    parser.add_argument('--full', action='store_true', help='Enable legacy Test.FullTest branches (may require interaction)')
    parser.add_argument('--clang', help='llvm-mingw clang executable for native Regex tests')
    parser.add_argument('--all', action='store_true', help='Include opt-in and unavailable suites')
    parser.add_argument('--list', action='store_true')
    parser.add_argument('--host', choices=['Excel', 'Word', 'Access'], default='Excel')
    parser.add_argument('--json', type=Path, help='Write structured results')
    args = parser.parse_args(argv)
    try:
        suites = load_manifest()
        if args.list:
            for suite in suites:
                if args.suite and suite['id'].casefold() not in {n.casefold() for n in args.suite}:
                    continue
                print(suite['id'] + ' [' + ', '.join(suite['hosts']) + '] ' + ('default' if suite['default'] else 'opt-in'))
                for case in suite['cases']:
                    print('  ' + case)
                if suite.get('unsupported') or suite.get('requirements'):
                    print('  Requires: ' + suite.get('unsupported', suite.get('requirements', '')))
            if args.suite:
                select(suites, args.suite, args.test, args.all)
            return 0
        selected = select(suites, args.suite, args.test, args.all)
    except ValueError as error:
        parser.error(str(error))
    results = []
    for suite, case in selected:
        start = time.monotonic()
        try:
            result = execute(suite, case, args.host, args.clang, args.full)
        except Exception as error:
            result = {'status': 'failed', 'detail': str(error)}
        result.update(suite=suite['id'], case=case, host=args.host, seconds=round(time.monotonic()-start, 3))
        results.append(result)
        print(result['status'].upper() + ' ' + suite['id'] + '.' + case + ': ' + result.get('detail', ''), flush=True)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_bytes((json.dumps({'results': results}, indent=2) + '\n').replace('\n', '\r\n').encode())
    counts = {status: sum(r['status'] == status for r in results) for status in ['passed', 'failed', 'skipped']}
    print(', '.join(str(count) + ' ' + status for status, count in counts.items()))
    return 1 if counts['failed'] or not counts['passed'] or (args.suite and counts['skipped']) else 0


if __name__ == '__main__':
    sys.exit(main())
