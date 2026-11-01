"""Opt-in integration tests for the shared Office runner's failure semantics."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'Tools/Testing'))
spec = importlib.util.spec_from_file_location('stdvba_runner', ROOT / 'Tools/Testing/run.py')
runner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(runner)


class OfficeLifecycleTests(unittest.TestCase):
    def execute(self, body):
        with tempfile.TemporaryDirectory(prefix='stdVBA-runner-contract-') as folder:
            path = Path(folder) / 'RunnerContract.bas'
            source = 'Attribute VB_Name = "RunnerContract"\nPublic Sub RunTest()\n' + body + '\nEnd Sub\n'
            path.write_bytes(source.replace('\n', '\r\n').encode())
            suite = dict(id='runner.contract', kind='vba', files=[str(path)],
                         cases={'Contract': 'RunnerContract.RunTest'}, hosts=['Excel'])
            return runner.execute(suite, 'Contract', 'Excel')

    def test_suppressed_assertion_failure_is_not_a_pass(self):
        result = self.execute('  On Error Resume Next\n  Test.Assert "deliberate failure", False')
        self.assertEqual(result['status'], 'failed')
        self.assertEqual(result['failures'], 1)
        self.assertIn('deliberate failure', result['detail'])

    def test_unexpected_error_is_not_a_pass(self):
        result = self.execute('  Err.Raise 5, , "deliberate runtime error"')
        self.assertEqual(result['status'], 'failed')
        self.assertIn('deliberate runtime error', result['detail'])

    def test_empty_case_is_not_a_pass(self):
        result = self.execute('  Dim unused As Long')
        self.assertEqual(result['status'], 'failed')
        self.assertIn('No assertions ran', result['detail'])


if __name__ == '__main__':
    unittest.main()
