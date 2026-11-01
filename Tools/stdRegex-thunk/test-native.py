"""Build real x86/x64 Windows test hosts and execute the embedded blobs."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile

HERE = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument('--clang', required=True)
args = parser.parse_args()
results = {}
with tempfile.TemporaryDirectory(prefix='stdRegex-abi-') as scratch:
    for arch, target, blob in [('x86', 'i686', 'matcher-x86.bin'), ('x64', 'x86_64', 'matcher.bin')]:
        exe = Path(scratch) / (arch+'.exe')
        subprocess.run([args.clang, '--target='+target+'-w64-windows-gnu', '-O2', '-static',
                        str(HERE/'test-native.c'), '-luser32', '-o', str(exe)], check=True)
        result = subprocess.check_output([str(exe), str(HERE/blob)], text=True, timeout=30).strip()
        assert '18000 checks passed' in result
        print(result)
        results[arch] = dict(result=result, blob_sha256=hashlib.sha256((HERE/blob).read_bytes()).hexdigest())
(HERE/'verification-native.json').write_bytes((json.dumps(results, indent=2)+'\n').replace('\n','\r\n').encode())
