# stdRegex integration

stdVBA owns the public `stdRegex` API, its compatibility behavior, tests, and the
tooling that packages it as one importable class. The regex engine remains in
[sihlfall/vba-regex](https://github.com/sihlfall/vba-regex).

## Source of truth

- Edit `Tools/stdRegex-build/stdRegex.wrapper.cls.in` for stdVBA API changes. It contains
  the wrapper only, with placeholders for the generated engine declarations and
  procedures.
- `Tools/stdRegex-build/engine.json` pins the upstream engine commit.
- `src/stdRegex.cls` is the generated distribution file. Import this file into VBA
  as usual; Python and the upstream checkout are only needed by maintainers.
- `tests/stdRegexTests.bas` tests the public contract, including compatibility
  comparisons against the previous list formatter.

Do not edit the generated class separately. There is one maintained wrapper, and
regeneration combines it with engine source from Git. The upstream `stdRegex3`
wrapper is independent; this build neither copies it nor invokes its class
generator. No upstream files are modified.

## Build and verify

Requires Python 3 and Git. With sibling `stdVBA` and `vba-regex` checkouts:

```powershell
python Tools/stdRegex-build/build.py
python Tools/stdRegex-build/build.py --check
```

For another upstream location, pass `--engine-root C:/path/to/vba-regex`.

### What engine.json does

`engine.json` is an **input to a normal build**, not an automatically generated
cache. It records the upstream repository and the exact commit used for the
distributed class. A default build reproduces that recorded version even if a
newer upstream release exists. That is intentional: rebuilding the same stdVBA
revision should not silently change its regex engine.

It is **not a permanent freeze**. The builder accepts any locally available
branch, tag, or commit:

```powershell
# Preview another engine version without changing the version record:
python Tools/stdRegex-build/build.py --engine-ref HEAD

# Adopt it: regenerate the class and update engine.json to its full commit SHA:
python Tools/stdRegex-build/build.py --engine-ref HEAD --update-engine
```

Here, `HEAD` means the engine checkout's current commit, not stdVBA's commit.
To select the newest upstream default branch without switching your checkout:

```powershell
git -C ../vba-regex fetch https://github.com/sihlfall/vba-regex.git HEAD
python Tools/stdRegex-build/build.py --engine-ref FETCH_HEAD --update-engine
```

The selected ref is resolved once to an immutable commit before reading engine
files. With `--update-engine`, the version record becomes an **output of that
explicit update**, written only after generation succeeds. Check in the updated
class and version record together after reviewing the diff and running the VBA
tests. A preview alone intentionally leaves the old record unchanged, so a plain
`--check` will detect a different version until it is explicitly adopted.

The builder itself does not fetch or switch branches. A future engine version
may require changes to the packaging adapter if upstream changes its module or
API structure; a generated class is not proof of runtime compatibility.

### Manual GitHub Actions build

The workflow is `.github/workflows/stdRegex_build.yml`, named **Build stdRegex**.
Once it is on the repository's default branch, open **Actions Ã¢â€ â€™ Build stdRegex Ã¢â€ â€™
Run workflow** and choose:

- **latest** (default): fetch the current upstream default branch, build it, and
  record its exact commit in the artifact's `engine.json`.
- **recorded**: rebuild the commit already recorded in the repository.
- **ref**: build the branch, tag, or commit entered in `engine_ref`.

The workflow runs the builder's version-selection tests and checks that the
generated class is reproducible. Its downloadable artifact contains
`src/stdRegex.cls`, `Tools/stdRegex-build/engine.json`, and `stdRegex-build.patch`.
The patch can be applied to the corresponding stdVBA revision with
`git apply stdRegex-build.patch`, then reviewed and tested before committing.
There is no automatic push to the repository. GitHub-hosted builds do not run
Office VBA tests; use the local runner before adopting an engine update.

Manual workflow details: [GitHub workflow_dispatch documentation](https://docs.github.com/en/actions/reference/workflows-and-actions/events-that-trigger-workflows#workflow_dispatch).

The packaging adaptations flatten module names, retain one shared copy of the
engine's static tables in the predeclared class, and preserve stdVBA's processing
limit error behavior. The matching algorithm is unchanged. License and commit
provenance are embedded in the distribution file.

On Windows with Excel or Word and **Trust access to the VBA project object model** enabled in the chosen host:

```powershell
./Tools/stdRegex-build/run-tests.ps1
# Or run the same VBA suite in Word:
./Tools/stdRegex-build/run-tests.ps1 -HostApplication Word
```

The runner checks that the distribution file is current, then imports the class
and tests into a scratch workbook or document in a separate hidden Office
instance. It closes the scratch file without saving and quits only that instance.
The small `TestReporter.cls` adapter counts assertions and reports failures
without requiring Excel worksheets. The runner does not change security settings.
Use the VBA editor's **Break on Unhandled Errors** mode so intentional error tests
can reach their handlers. Tests can also be run using the usual test workbook.

## Benchmarks

The [measured before/after results](benchmarks/list-and-matchall/results.md) include
all samples, environment details, and the baseline commit. To reproduce:

```powershell
./Tools/stdRegex-build/benchmark.ps1 -HostApplication Word
```

The runner loads the pre-change pure-VBA class and current class into the same
hidden Office process. It checks output equality, warms each path, calibrates
batches, rotates operation order, and measures seven samples per case using
`QueryPerformanceCounter`. Pattern compilation is excluded; returned values are
allocated and released inside the measured interval. Medians, minimums, and
maximums are recorded, with all raw samples in JSON and CSV.

The six cases cover no matches, one match, 100 matches, 1,000 matches, no captures,
and sparse matches. `ListArr` returns two formatted columns. `MatchAllArr` and
`MatchAll` return different structures, so their comparison measures the cost of
each API's contract, not identical return objects. In particular, one names
dictionary per `MatchAllArr` call can make it slower when there are no matches.

Use `-OutputDir <path>`, `-BaselineRef <commit>`, `-Samples <count>`, and
`-TargetSeconds <minimum-batch-duration>` to control a run. Timings describe the
measured host and workloads, not a guaranteed speedup for every input.

## List and ListArr

Both methods scan all matches regardless of the `g` flag and format capture
positions directly. They do not call `MatchAll` or create match dictionaries.
`List` builds its output with a string builder. `ListArr` retains its 1-based row
and column bounds, Variant cells containing strings, and `Array()` empty result.

These methods intentionally retain their legacy formatting rules, including
`$&`, `$0`, numbered and `$name` captures, `\r`, `\n`, and `\$`. They also retain
the old sequential replacement order, including replacements within inserted
capture text and overlapping tokens such as `$1` / `$10`. This is different from
`Replace`'s engine format syntax; the build does not silently migrate list
formats. `MatchAll` continues returning a Collection of Dictionaries.

## MatchAllArr

```vb
Dim rx As stdRegex, values As Variant, names As Object
Set rx = stdRegex.Create("abc(?<hello>.*?)efg")
values = rx.MatchAllArr("abcdefg and abc something efg", names)

' values(1, 0) = "abcdefg"
' values(1, 1) = "d"
' values(2, 0) = "abc something efg"
' values(2, 1) = " something "
' names("hello") = 1
```

The signature is `MatchAllArr(ByVal sHaystack As String, Optional ByRef captureNames As Object) As Variant`.

- Rows run from 1 through the match count.
- Columns run from 0 through the number of capturing groups; column 0 is the whole
  match. Each cell is a string. Noncapturing groups do not add columns.
- Unmatched captures, including declarations removed by a `{0}` quantifier, contain `""`. With no captures, column 0 remains.
- With no matches, the result is `Array()`, consistent with `ListArr`.
- All successive, nonoverlapping matches are returned regardless of `g`. Other
  matching flags and processing-limit errors behave as they do in `MatchAll`.
- `captureNames` receives one new, case-sensitive dictionary for the entire call,
  including names of groups that did not match or when no matches were found.
  A supplied previous object is replaced. Duplicate names map to their first
  declared numbered capture, even when a later group participates in a match.
- The output argument may be omitted: `values = rx.MatchAllArr(text)`.

The example uses `.*?` because `\w+` does not capture spaces.

## Optional native matching (WIP)

The generated class includes both the upstream VBA matcher and stdVBA's native
x86/x64 port. Native is the default. Opt into the debuggable VBA matcher with:

```vb
Set rx = stdRegex.Create("\d+", useNative:=False)
```

`useNative` is the fourth optional constructor argument, after `stepsLimit`.
The constructor cache distinguishes engine selection; changing Pattern or Flags
preserves it. Native matching requires Windows VBA7, with automatic Office-bitness
selection. Compilation and public result formatting remain VBA. Native memory
is allocated only when native matching is first used.

`../stdRegex-thunk/build.py --clang <path>` rebuilds the owned C matcher blobs.
The normal class build embeds the checked-in blobs without requiring a compiler.
It verifies blob hashes and the upstream matching interpreter fingerprint;
upstream matcher changes require a reviewed native-port update before generation
can continue. Neither this port nor its build changes are sent upstream.

Run `python Tools/stdRegex-thunk/run.py --test-only` for both engines' public
contract suites and differential comparisons in one disposable Word process.
See `../stdRegex-thunk/README.md` for native ABI tests and platform limitations.
