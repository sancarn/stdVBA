"""Build relocation-free x86 and x64 matching thunks and embed it in a standalone class."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import struct
import subprocess
import tempfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent

def write(path, text):
    path.write_bytes(text.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8'))

def build_code(clang, target, expected_machine, symbol):
    flags = ['--target='+target, '-O2', '-c', '-ffreestanding',
             '-fno-builtin', '-fno-stack-protector', '-fno-jump-tables',
             '-fno-vectorize', '-fno-slp-vectorize', '-fno-asynchronous-unwind-tables']
    with tempfile.TemporaryDirectory(prefix='stdRegex-native-build-') as temp:
        obj = Path(temp) / 'matcher.o'
        subprocess.run([clang, *flags, str(HERE/'matcher.c'), '-o', str(obj)], check=True)
        data = obj.read_bytes()
    machine, sections, _, symptr, nsyms, optional, _ = struct.unpack_from('<HHIIIHH', data)
    assert machine == expected_machine and optional == 0
    code = None
    for index in range(sections):
        off = 20 + 40*index
        name = data[off:off+8].rstrip(b'\0')
        size, pointer, reloc, _, nreloc = struct.unpack_from('<IIIIH', data, off+16)
        if name == b'.text':
            assert nreloc == 0, 'Native thunk must have no relocations/imports'
            code = data[pointer:pointer+size]
            text_section = index+1
    assert code
    found = False
    for idx in range(nsyms):
        off = symptr + idx*18
        name = data[off:off+8]
        if name[:4] == b'\0'*4:
            pos = symptr+nsyms*18+struct.unpack_from('<I', name, 4)[0]
            name = data[pos:data.index(b'\0', pos)]
        else:
            name = name.rstrip(b'\0')
        if name == symbol:
            value, section = struct.unpack_from('<Ih', data, off+8)
            assert value == 0 and section == text_section, 'Entry must be first in .text'
            found = True
    assert found
    return code, flags

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--clang', required=True)
    args = parser.parse_args()
    builds = {name: build_code(args.clang, *config) for name, config in {
        'x64': ('x86_64-w64-windows-gnu', 0x8664, b'regex_entry'),
        'x86': ('i686-w64-windows-gnu', 0x14c, b'_regex_entry@16'),
    }.items()}
    metadata = {}
    for arch, (code, flags) in builds.items():
        (HERE/('matcher.bin' if arch == 'x64' else 'matcher-x86.bin')).write_bytes(code)
        metadata[arch] = dict(native_bytes=len(code), native_sha256=hashlib.sha256(code).hexdigest(), flags=flags, relocations=0)
        print(f'{arch}: embedded {len(code)} bytes, zero relocations')
    write(HERE/'build.json', json.dumps({'engine_matcher_sha256':json.loads((HERE/'build.json').read_text())['engine_matcher_sha256'],
        'architectures':metadata,
        'compiler':subprocess.check_output([args.clang,'--version'],text=True).splitlines()[0]}, indent=2)+'\n')

if __name__ == '__main__':
    main()
