"""Version-selection and failure-safety tests for the standalone-class builder."""
from contextlib import redirect_stderr, redirect_stdout
import io
import json
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import build


class BuildSelectionTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.here = root / 'stdVBA' / 'Tools' / 'stdRegex-build'
        self.here.mkdir(parents=True)
        self.target = root / 'stdVBA' / 'src' / 'stdRegex.cls'
        self.target.parent.mkdir()
        self.target.write_bytes(b'previous class\r\n')
        self.engine = root / 'engine'
        self.engine.mkdir()
        self.git('init', '-q')
        self.git('config', 'user.name', 'Builder test')
        self.git('config', 'user.email', 'test@example.invalid')
        (self.engine / 'LICENSE').write_text('first engine revision')
        self.git('add', 'LICENSE')
        self.git('commit', '-q', '-m', 'first')
        self.first = self.git('rev-parse', 'HEAD').strip()
        (self.engine / 'LICENSE').write_text('second engine revision')
        self.git('commit', '-qam', 'second')
        self.second = self.git('rev-parse', 'HEAD').strip()
        self.manifest = self.here / 'engine.json'
        self.manifest.write_text(json.dumps({'repository': 'https://github.com/sihlfall/vba-regex', 'revision': self.first}) + '\n')
        self.saved_manifest = self.manifest.read_bytes()
        self.saved_class = self.target.read_bytes()

    def git(self, *args):
        return subprocess.check_output(['git', '-C', str(self.engine), *args], text=True)

    def invoke(self, *args, fail=False):
        def generate(root, selected):
            if fail:
                raise ValueError('incompatible engine')
            return 'generated ' + selected['revision'] + '\n'
        with patch.object(build, 'HERE', self.here), patch.object(build, 'generate', side_effect=generate), redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            build.main(['--engine-root', str(self.engine), *args])

    def test_default_reproduces_recorded_commit_even_when_head_is_newer(self):
        self.invoke()
        self.assertEqual(self.target.read_text(), 'generated ' + self.first + '\n')
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)

    def test_ref_override_does_not_adopt_revision_implicitly(self):
        self.invoke('--engine-ref', 'HEAD')
        self.assertEqual(self.target.read_text(), 'generated ' + self.second + '\n')
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)

    def test_update_records_resolved_sha_and_matching_class(self):
        self.invoke('--engine-ref', 'HEAD', '--update-engine')
        self.assertEqual(json.loads(self.manifest.read_text())['revision'], self.second)
        self.assertEqual(self.target.read_bytes(), ('generated ' + self.second + '\r\n').encode())

    def test_update_requires_a_ref(self):
        with self.assertRaises(SystemExit):
            self.invoke('--update-engine')
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)
        self.assertEqual(self.target.read_bytes(), self.saved_class)

    def test_check_is_read_only(self):
        self.target.write_text('generated ' + self.first + '\n')
        before = self.target.read_bytes()
        self.invoke('--check')
        self.assertEqual(self.target.read_bytes(), before)
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)

    def test_check_cannot_update_manifest(self):
        with self.assertRaises(SystemExit):
            self.invoke('--check', '--engine-ref', 'HEAD', '--update-engine')
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)

    def test_incompatible_engine_does_not_change_existing_files(self):
        with self.assertRaises(ValueError):
            self.invoke('--engine-ref', 'HEAD', '--update-engine', fail=True)
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)
        self.assertEqual(self.target.read_bytes(), self.saved_class)

    def test_missing_revision_does_not_change_existing_files(self):
        with self.assertRaises(subprocess.CalledProcessError):
            self.invoke('--engine-ref', 'missing-engine-ref', '--update-engine')
        self.assertEqual(self.manifest.read_bytes(), self.saved_manifest)
        self.assertEqual(self.target.read_bytes(), self.saved_class)


if __name__ == '__main__':
    unittest.main()
