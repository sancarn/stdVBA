"""Contract checks for suite selection and generated VBA bridges; no Office needed."""
import importlib.util
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('stdvba_runner', ROOT / 'Tools/Testing/run.py')
runner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(runner)


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.suites = runner.load_manifest()

    def test_every_root_test_module_registered(self):
        registered = {f for s in self.suites for f in s.get('files', [])}
        adapter_modules = {'stdImageTests', 'stdHTTPLocalTests', 'stdWebSocketLocalTests',
                           'stdWebSocketAsyncTests', 'stdWindowSubclassTests'}
        for path in (ROOT / 'tests').glob('*Tests.bas'):
            self.assertTrue(path.relative_to(ROOT).as_posix() in registered or path.stem in adapter_modules, path)

    def test_individual_case_does_not_select_siblings(self):
        selected = runner.select(self.suites, ['STDCALLBACK'], 'bind')
        self.assertEqual([(s['id'], c) for s, c in selected], [('stdCallback', 'Bind')])

    def test_unknown_selectors_fail(self):
        for names, case in [(['missing'], None), (['stdCallback'], 'missing'), (None, 'Bind')]:
            with self.assertRaises(ValueError):
                runner.select(self.suites, names, case)

    def test_all_includes_opt_in_and_unavailable_suites(self):
        selected = {s['id'] for s, c in runner.select(self.suites, all_suites=True)}
        self.assertEqual(selected, {s['id'] for s in self.suites})

    def test_bridge_checks_failures_and_zero_assertions(self):
        suite, case = runner.select(self.suites, ['stdCallback'], 'Bind')[0]
        source = runner.bridge(suite, case)
        self.assertIn('Call stdCallbackTests.TestBind', source)
        self.assertIn('Test.failures = 0', source)
        self.assertIn('No assertions ran', source)
        self.assertNotIn('On Error Resume Next', source)

    def test_return_contracts_are_checked(self):
        for name, expected in [('stdJSON', 'Left$(CStr(result), 4)'), ('stdRefArray', 'result > 0'), ('stdTimer', 'Left$(CStr(result), 4)')]:
            suite, case = runner.select(self.suites, [name])[0]
            self.assertIn(expected, runner.bridge(suite, case))


if __name__ == '__main__':
    unittest.main()
