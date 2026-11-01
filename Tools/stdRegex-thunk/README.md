# Native regex matching experiment (WIP)

The normal generated `src/stdRegex.cls` now contains both matching engines.
Use `stdRegex.Create(pattern, flags, stepsLimit, useNative:=False)` to select VBA; omitted/True selects Windows x86/x64 native matching. Existing
argument positions are preserved; the default engine changes to native. Selection is part of the cache
key and follows the compiled regex, including Pattern/Flags changes.

Both machine-code blobs are embedded automatically by
`Tools/stdRegex-build/build.py`; importing the class requires no DLL, compiler
or helper module. Upstream sihlfall/vba-regex is not modified. An upstream matcher
change fails generation until its native port has been reviewed and updated.

The complete DFS matching interpreter is ported to `matcher.c`, compiled to
one architecture-specific, relocation-free machine-code block: 5,383 bytes for
x64 and 5,190 bytes for x86. Both blobs are embedded in the same class;
`#If Win64 Then` selects the correct one for Office bitness automatically.
Matching includes the candidate search loop, all existing interpreter opcodes,
backtracking, capture copy-on-write, named captures, backreferences,
lookahead/lookbehind, bounded quantifiers, flags and per-candidate step limits.
The original VBA parser/compiler, Unicode canonicalization data and public
result formatting remain VBA. This is a native **matcher**, not native pattern
compilation. No callback into VBA occurs inside the machine-code matcher.

Executable memory is allocated writable, copied, then changed to read/execute;
the instruction cache is flushed. The predeclared singleton owns that block
and reusable workspace. Normal class termination frees the executable block.
VBE Reset/End bypasses cleanup and can leak that allocation; this remains an
experimental class. There is no Mac or pre-VBA7 implementation or silent fallback.

## Reproduce

The build used portable LLVM-MinGW 20260602 (Clang 22.1.7), available from
https://github.com/mstorsjo/llvm-mingw/releases/tag/20260602 .
The build tool inspects the COFF object and rejects text relocations or a
nonzero entry offset, rather than embedding a DLL or unresolved object file.

```powershell
python Tools/stdRegex-thunk/build.py --clang C:/path/to/llvm-mingw/bin/clang.exe
python Tools/stdRegex-build/build.py --engine-root C:/path/to/vba-regex
python Tools/stdRegex-thunk/run.py
python Tools/stdRegex-thunk/summarize.py
```

Testing requires desktop Word, Python, pywin32 and psutil. The runner launches
a disposable hidden Word instance, checks imported files' actual CRLF bytes,
temporarily configures the installed Word version's AccessVBOM and VBE error
trapping, and restores the original registry values/types/existence on exit.
A separate supervisor dismisses and records modal dialogs only in its owned
Office process, minimizes visible owned windows, and imposes a 240-second
deadline. Any detected dialog fails the run. Managed access restrictions are
respected. Existing user Office processes are not attached to or terminated.

## Validation and measurement

The existing public contract suite runs against both baseline and native classes.
Additional differential cases cover empty matches, greediness, nested captures,
backreferences, named groups, lookarounds, Unicode/embedded NUL input, local and
global flags, workspace growth and processing-limit errors. These checks are
regression evidence, not proof for every pattern or host.

`results/results.json` contains raw samples, assertion counts, baseline commit,
class hash, host details and dialog status. `results/summary.json` contains
medians, minimums, maximums and ratios. Open `results/comparison.html` to switch
between the four measured operations and inspect the chart and timings.

Both implementations run in the same Word process. Patterns are cached and
warmed before timing; initialization, native allocation and compilation are
excluded. Each operation calibrates a batch to at least 50 ms, then records
seven samples with rotating execution order using QueryPerformanceCounter.
`Test`, `List`, `ListArr`, and `MatchAllArr` are compared against the **same API**
on the baseline. Result allocation/release is included. `Test` stops at the
first match, so increasing the number of matches does not make it scan all of
them. Differential checks precede all timing.

The large scan gains do not remove VBA/COM formatting costs. In particular,
creating the names dictionary can dominate small `MatchAllArr` calls. No
compilation speedup is claimed. Measurements apply to this host and workloads.

## Compatibility review

`Create` and protected `protInit` append optional `useNative As Boolean = True`
after `stepsLimit`. Existing calls now use native matching; pass `useNative:=False` to retain VBA.
Native mode uses the same pattern compiler, capture layout, comparison rules,
public result types and step-limit errors. Pattern/Flags setters preserve the
selected engine. Unsupported hosts reject an explicit native request with error
5. The existing Dictionary dependency still prevents claiming Mac support.
VBA6/Mac compilation excludes native declarations and code.

The native port is experimental. Full 32-bit Office execution remains unverified;
see native ABI validation below. VBE Reset/End can leak the executable allocation.

The C port derives from sihlfall/vba-regex revision
`1619cd75526903045b8c5cb8c6048f93bc4160be`; see `LICENSE.engine` and the retained
license in the generated class. `build.json` records compiler flags and hashes.

## 32-bit validation

The exact extracted x86 and x64 blobs each passed 18,000 checks in separately
compiled Windows executables through `CallWindowProcW`. The x86 executable runs
under WoW64 and exercises the 32-bit stdcall ABI, repeated stack cleanup, the
52-byte argument layout (72 bytes on x64), captures, branches, backreferences,
flags, scanning, no-match and negative error returns. Both object files have
zero code relocations. This checks executable code, not just cross-compilation.

`test-native.c` contains 18 fixture checks repeated 1,000 times per architecture.
Build with the same Clang using `--target=i686-w64-windows-gnu -O2 -static
Tools/stdRegex-thunk/test-native.c -luser32 -o <scratch>/test-x86.exe`, then pass
`Tools/stdRegex-thunk/matcher-x86.bin` to that executable. Repeat with
`--target=x86_64-w64-windows-gnu` and `matcher.bin` for 64-bit.

Only 64-bit Office is installed on the development host. The generated class with both engines passed the existing 328 assertions and 2,248 differential cases in Word
x64 (`python Tools/stdRegex-thunk/run.py --test-only`; results in
`verification-x64`). Full 32-bit Office import/execution remains unverified;
no 32-bit benchmark speedup is claimed. Historical benchmark files under `results` record the earlier standalone
experimental class; `verification-x64` validates the current integrated class.
