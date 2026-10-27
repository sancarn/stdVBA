# Building stdRegex

Import `src/stdRegex.cls` as usual. Its regex engine is bundled into the class;
no additional VBA modules, references, DLLs, or ASF framework classes are needed.
The existing late-bound `Scripting.Dictionary` dependency remains for the cache
and match results, so this change does not make the wrapper Mac-compatible.

The engine comes from [sihlfall/vba-regex](https://github.com/sihlfall/vba-regex/tree/1619cd75526903045b8c5cb8c6048f93bc4160be),
pinned to commit `1619cd75526903045b8c5cb8c6048f93bc4160be`. Its MIT license is
included in the generated class, preserving single-file redistribution.

Edit `Wrapper.cls`, then run `./Tools/stdRegex/Build.ps1` from the repository root.
The build downloads the pinned sources into a temporary directory, runs the upstream
bundlers, and combines the engine with our compatibility wrapper. To build from an
existing checkout of that revision, pass `-UpstreamRoot <checkout-path>`.
Do not hand-edit the generated `src/stdRegex.cls`.

## Compatibility

`Create`, `Pattern`, `Flags`, `Test`, `Match`, `MatchAll`, `Replace`, `List`, and
`ListArr` retain their existing calling conventions. `Create` still caches compiled
instances. Match dictionaries retain numbered/named captures and `$COUNT`.
`MatchAll`, `List`, and `ListArr` consider all matches regardless of `g`; `Replace`
uses `g` to choose first/all. Legacy `$name` replacements and list escape sequences
remain supported. Replacement also accepts the engine's `$<name>` syntax.

Flags `i`, `g`, `m`, `x`, `c`, and `s` remain available. `x` removes unescaped ASCII
spaces; `c` removes line comments beginning with an unescaped `#`. Both preserve
character-class contents. `s` uses the engine's native dot-all mode. Flag spelling
and casing retain the old wrapper's behavior.

There are deliberate changes:

- `$RAW` is removed. It exposed a VBScript `IMatchCollection2`, which cannot be
  returned by a pure VBA engine. Read numbered/named captures instead.
- `Create(pattern, flags, stepsLimit)` adds an optional processing budget, default
  10000 steps per candidate match. Exhaustion raises `vbObjectError + 3051` rather
  than returning a false negative. Increase the budget for complex valid patterns.
- `ListArr` with no matches returns `Array()` instead of attempting an invalid
  zero-row allocation.
- Patterns now follow the upstream engine's syntax. Invalid patterns fail at
  construction or assignment. A failed pattern assignment preserves the old state.
  This is tested compatibility for the supported wrapper API, not a claim of full
  VBScript syntax/Unicode equivalence. 32-bit Office and Mac were not tested.

The integration uses native named captures rather than stripping names with
VBScript regexes. The build also fixes upstream standalone guard calls to use
`Err.Raise`, initializes new instances' budgets, and makes replacement iteration
raise when its budget is exhausted.

## Why this engine

Sihlfall already provides a `stdRegex3` wrapper and a standalone-class bundler.
The integration work is adapting that bundle to the existing stdVBA wrapper,
especially its six flags, `$name` formats, `ListArr`, and cached construction.
The upstream wrapper alone is not a compatible file rename: its list API,
replacement names, flags, and caching behavior differ.

[W. Garcia's ASF_RegexEngine](https://github.com/ECP-Solutions/ASF/blob/290ad3eabff97fc16f34568d83139994f89b83ac/src/ASF_RegexEngine.cls)
is also a standalone class. It uses built-in VBA Collections and does not require
ASF's compiler, virtual machine, maps, DLLs, or external references. Its `Init`,
`Test`, `Exec`, `ExecAll`, and named-capture results could be wrapped, but its
current `ParseEscapeNode` handles neither numeric backreferences nor word
boundaries: `\1` becomes literal `1` and `\b` becomes literal `b`. The
[upstream tests](https://github.com/ECP-Solutions/ASF/blob/290ad3eabff97fc16f34568d83139994f89b83ac/test/RegexTests.bas)
explicitly classify pattern backreferences as unsupported. Replacing stdRegex
with it would therefore require engine changes before it could pass issue 141
or retain existing pattern semantics. Dependencies are not the obstacle.

## Performance benchmark

The benchmark covers literal scanning, short validation, repeated captures,
case-insensitive alternation, multiline records, no-match scanning, word boundaries,
and a numeric backreference. Each workload measures fresh construction plus first
Test, reused Test, MatchAll, and global Replace. Two isolated Excel runs produce
ten timed batches per cell. The scripts save raw samples and machine/source metadata
alongside a summary and can generate a standalone HTML report.

To reproduce from the repository root:

```powershell
./Tools/stdRegex/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdRegex/benchmarks/latest
# Run again to capture a second independent sample set:
./Tools/stdRegex/Benchmark.ps1 -TemporaryTrustAccess -OutputDirectory ./Tools/stdRegex/benchmarks/latest/repeat
python ./Tools/stdRegex/Analyze.py ./Tools/stdRegex/benchmarks/latest
python ./Tools/stdRegex/BuildReport.py ./Tools/stdRegex/benchmarks/latest
```

The old class is read from pinned commit `abf14dd81e553cbcf438284975c80f3e9220196a`
and renamed only to coexist with the migrated class. The standalone ASF class is
downloaded at the revision linked above; `-AsfSource` accepts a local copy of that
same revision. Neither source is modified to fix behavior or improve speed.

Timing uses QueryPerformanceCounter inside VBA, excluding Excel automation calls.
Each cell checks the expected Boolean, all numbered capture values/counts, or full
replacement output as appropriate, then calibrates to at least 30 ms per batch
(at most 16384 iterations). Five measured batches follow; engine order rotates
between cells. Summary values are medians of per-call batch averages, not individual
latency percentiles. Construction bypasses Create caches; other operations reuse
preinitialized regex objects. Shared static initialization is warmed up first.

Native processing budgets are raised to 2147483647 to prevent repeated replacements
from hitting ASF's accumulated step count. All chosen inputs are small and do not
exercise pathological backtracking. The compatibility suite runs
before every benchmark. Excel access and debugger settings are restored afterward.
Results are specific to this machine and workload; ranges and separate run medians
in the summary show noise from background load and COM allocation.

## Tests

Run `./Tools/stdRegex/Test.ps1` with desktop Excel installed and VBA project access
enabled. `-TemporaryTrustAccess` explicitly enables VBA project access for the run
and restores the previous value afterward. The runner opens a separate Excel
process, imports only `stdRegex`, its tests, and an assertion module, then discards
the workbook. It temporarily selects unhandled-error trapping for expected-error
tests and restores the debugger preferences afterward.

The suite includes the exact issue 141 reproducer `(.)(?=.*\1)` against `ss` and
the issue 162 dot-all constructor, plus the legacy API, flags, setters,
backreferences, word boundaries, empty matches, replacement formats, and
processing-budget errors.
