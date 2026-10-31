"""Regenerate the embedded subclass thunks. Dev dependency: keystone-engine."""
import argparse
from pathlib import Path
import re
from keystone import Ks, KS_ARCH_X86, KS_MODE_32, KS_MODE_64

ROOT = Path(__file__).resolve().parents[2]


def assemble(bits):
    source = Path(__file__).with_name('thunk%d.asm' % bits).read_text()
    source = '\n'.join(line.split(';')[0] for line in source.splitlines())
    code = bytes(Ks(KS_ARCH_X86, KS_MODE_64 if bits == 64 else KS_MODE_32).asm(source)[0])
    assert len(code) < 3072, 'Code overlaps the unwind table'
    if bits == 64:
        assert code.startswith(bytes.fromhex('4881ECD8000000')), 'Update the VBA unwind record for the new prologue'
    return code


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    path = ROOT / 'src/stdWindow.cls'
    source = path.read_bytes().decode('utf-8')
    lines = ["  'BEGIN GENERATED SUBCLASS THUNKS", '#If Win64 Then']
    for bits in (64, 32):
        if bits == 32:
            lines.append('#Else')
        code = assemble(bits).hex().upper()
        lines.append('  sCode = "' + code[:96] + '"')
        for start in range(96, len(code), 96):
            lines.append('  sCode = sCode & "' + code[start:start+96] + '"')
    lines.extend(['#End If', "  'END GENERATED SUBCLASS THUNKS"])
    updated, count = re.subn(r"  'BEGIN GENERATED SUBCLASS THUNKS.*?  'END GENERATED SUBCLASS THUNKS",
                             '\r\n'.join(lines), source, flags=re.S)
    assert count == 1
    if args.check:
        assert updated == source, 'Embedded thunks need regeneration'
    else:
        path.write_bytes(updated.encode('utf-8'))


if __name__ == '__main__':
    main()
