"""Embed stdVBA-owned optional native matching in the upstream VBA engine."""
import hashlib
import json
from pathlib import Path
HERE = Path(__file__).resolve().parent

def integrate(source):
    start = source.index('Private Function DfsMatch(')
    end = source.index('Private Sub ParseFormatString(', start)
    metadata = json.loads((HERE/'build.json').read_text())
    actual = hashlib.sha256(' '.join(source[start:end].split()).encode()).hexdigest()
    if actual != metadata['engine_matcher_sha256']:
        raise ValueError('Upstream matcher changed: review/port the native engine and update its fingerprint before building')
    template = (HERE/'native.bas.in').read_text()
    declarations, procedures = template.split("' @PROCEDURES@")
    for arch, blob in [('x64','matcher.bin'),('x86','matcher-x86.bin')]:
        data = (HERE/blob).read_bytes()
        if hashlib.sha256(data).hexdigest() != metadata['architectures'][arch]['native_sha256']:
            raise ValueError('Native blob hash mismatch: '+arch)
        hexcode = data.hex().upper()
        chunks = '\n'.join('    hexCode = hexCode & "'+hexcode[i:i+120]+'"' for i in range(0,len(hexcode),120))
        procedures = procedures.replace("    ' @HEX_"+arch.upper()+"@", chunks)
    source = source.replace('Option Explicit', 'Option Explicit\n#If VBA7 And Not Mac Then\n'+declarations+'\n#End If', 1)
    source = source.replace('Private Type TThis\n', 'Private Type TThis\n#If VBA7 And Not Mac Then\n  nativeCode As LongPtr\n  nativeCapacity As Long\n  nativeWork() As Long\n#End If\n',1)
    source = source.replace('Private Type RegexTy\n', 'Private Type RegexTy\n    useNative As Boolean\n',1)
    # Selection travels with each compiled regex, never with shared singleton state.
    for name in ['DfsMatch','DfsMatchFrom']:
        a=source.index('Private Function '+name+'(')
        b=source.index(') As Long\n',a)
        source=source[:b].rstrip().removesuffix('_').rstrip()+', _\n    Optional ByVal useNative As Boolean = False _\n'+source[b:]
    source=source.replace('dotAll:=dotAll _\n    )', 'dotAll:=dotAll, useNative:=useNative _\n    )',1)
    source=source.replace('stepsLimit:=regex.stepsLimit, _', 'stepsLimit:=regex.stepsLimit, useNative:=regex.useNative, _')
    a=source.index('Private Function DfsMatchFrom(')
    b=source.index(') As Long\n',a)+len(') As Long\n')
    branch='#If VBA7 And Not Mac Then\n    If useNative Then\n        DfsMatchFrom = NativeMatch(outCaptures, bytecode, inputStr, sp, stepsLimit, multiline, dotAll, raiseOnLimit)\n        Exit Function\n    End If\n#End If\n'
    source=source[:b]+branch+source[b:]
    return source+'\n#If VBA7 And Not Mac Then\n'+procedures.replace('stdRegexThunk','stdRegex')+'\n#End If\n'
