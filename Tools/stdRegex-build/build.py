"""Build stdVBA's standalone stdRegex from its wrapper and a recorded or selected engine commit."""

import argparse
import json
from pathlib import Path
import re
import subprocess


HERE = Path(__file__).resolve().parent
MODULES = (
    "StaticStringBuilder", "ArrayBuffer", "RegexNumericConstants", "RegexErrors",
    "RegexBytecode", "RegexAst", "RegexUnicodeSupport", "RegexRanges",
    "RegexIdentifierSupport", "RegexLexer", "RegexCompiler", "RegexDfsMatcher",
    "RegexReplace", "StaticRegex",
)
OMIT = {
    "UnicodeInitialize", "RangeTablesInitialize", "AstTableInitialize",
    "TryInitializeRegex", "GetCapture", "GetCaptureByName",
    "InitializeMatcherState", "ResetMatcherState",
}
PROC = re.compile(
    r"^(?:Public|Private|Friend) (Sub|Function) (\w+)\b.*?^End \1[ \t]*(?:\n|$)",
    re.M | re.S,
)


def read_revision(root, revision, name):
    return subprocess.check_output(
        ["git", "-C", str(root), "show", f"{revision}:{name}"],
    ).decode("utf-8-sig").replace("\r\n", "\n")


def unqualify(text):
    for old, new in (
        ("StaticStringBuilder.Ty", "StaticStringBuilder"),
        ("ArrayBuffer.Ty", "ArrayBuffer"), ("RegexLexer.Ty", "LexerContext"),
    ):
        text = text.replace(old, new)
    for module in MODULES:
        text = text.replace(module + ".", "")
    return text


def engine_parts(root, revision):
    declarations, procedures = [], []
    for module in MODULES:
        text = read_revision(root, revision, f"src/{module}.bas")
        own_type = {"StaticStringBuilder": "StaticStringBuilder", "ArrayBuffer": "ArrayBuffer", "RegexLexer": "LexerContext"}.get(module)
        if own_type:
            text = re.sub(r"(?<!\.)\bTy\b", own_type, text)
        text = re.sub(r"^(?:Attribute |Option |').*(?:\n|$)", "", text, flags=re.M)
        first = PROC.search(text)
        if first is None:
            declarations.append(unqualify(re.sub(r"^Public\b", "Private", text, flags=re.M)))
            continue
        decl, body = text[:first.start()], text[first.start():]
        # The predeclared class stores the engine's static tables. Instances only
        # store their own compiled pattern; do not duplicate the large tables.
        decl = re.sub(
            r"^(?:Public|Private|Dim) (?!(?:Const|Type|Enum)\b).*\n", "", decl, flags=re.M,
        )
        if module == "RegexDfsMatcher":
            decl = re.sub(r"^Public Type StartLengthPair\n.*?^End Type\n", "", decl, flags=re.M | re.S)
        decl = re.sub(r"^Public\b", "Private", decl, flags=re.M)
        declarations.append(unqualify(decl))

        static_names = re.findall(r"^Public (?:Sub|Function) (\w+)\b", body, re.M) if module == "StaticRegex" else []

        def adapt(match):
            kind, name = match.group(1, 2)
            if name in OMIT:
                return ""
            proc = unqualify(match.group())
            proc = re.sub(r"^Public\b", "Private", proc)
            proc = re.sub(
                r"^.*If Not (?:astTableInitialized|UnicodeInitialized|RangeTablesInitialized) Then.*\n",
                "", proc, flags=re.M,
            )
            proc = re.sub(r"\bStaticData\b", "This.regex.bytecode", proc)
            if name == "DfsRunThreads":
                proc = proc.replace(
                    "If raiseOnLimit Then Err.Raise REGEX_ERR_STEP_LIMIT",
                    'If raiseOnLimit Then Err.Raise REGEX_ERR_STEP_LIMIT, "stdRegex", "Regex processing step limit exceeded"',
                )
            if module == "StaticRegex":
                for public_name in static_names:
                    proc = re.sub(r"\b" + public_name + r"\b", "protStaticRegex" + public_name, proc)
                proc = re.sub(r"^Private\b", "Friend", proc)
                # Convenience operations must propagate the same processing-limit
                # errors as the public stdVBA methods.
                proc = proc.replace(
                    "protStaticRegexMatchNext(matcherState, regex, haystack)",
                    "protStaticRegexMatchNext(matcherState, regex, haystack, raiseOnLimit:=True)",
                )
                lines = proc.splitlines()
                end = 0
                while lines[end].rstrip().endswith("_"):
                    end += 1
                lines[end + 1:end + 1] = [
                    "    If Not (Me Is stdRegex) Then",
                    '        Err.Raise 1, "stdRegex", "Method called on object, not on class"',
                    f"        Exit {kind}", "    End If", "",
                ]
                proc = "\n".join(lines) + "\n"
            return proc

        procedures.append(PROC.sub(adapt, body))
    return "\n".join(declarations), "\n".join(procedures)


def resolve_revision(root, ref):
    """Resolve once so moving branches cannot change halfway through a build."""
    return subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", "--verify", "--end-of-options", ref + "^{commit}"],
        text=True,
    ).strip()


def generate(root, pin):
    revision = pin["revision"]
    decl, body = engine_parts(root, revision)
    license_text = read_revision(root, revision, "LICENSE")
    provenance = "\n".join([
        "' Generated by Tools/stdRegex-build/build.py; edit stdRegex.wrapper.cls.in instead.",
        "' Bundled " + pin["repository"] + " " + revision,
        *("' " + line if line else "'" for line in license_text.splitlines()), "",
    ])
    template = (HERE / "stdRegex.wrapper.cls.in").read_text(encoding="utf-8")
    for marker, value in (("' @ENGINE_DECLARATIONS@", provenance + decl), ("' @ENGINE_PROCEDURES@", body)):
        if template.count(marker) != 1:
            raise ValueError(f"Expected exactly one {marker} marker")
        template = template.replace(marker, value.rstrip())
    import importlib.util
    spec = importlib.util.spec_from_file_location("stdregex_native", Path(__file__).resolve().parent.parent / "stdRegex-thunk" / "integrate.py")
    native = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(native)
    template = native.integrate(template)
    template = "\n".join(line.rstrip() for line in template.splitlines())
    return re.sub(r"\n(?:[ \t]*\n){2,}", "\n\n", template).rstrip() + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine-root", type=Path, default=HERE.parents[2] / "vba-regex")
    parser.add_argument("--check", action="store_true", help="Fail if the committed class needs regeneration")
    parser.add_argument("--engine-ref", help="Use this local engine branch, tag, or commit instead of engine.json")
    parser.add_argument("--update-engine", action="store_true", help="Record the selected exact commit in engine.json after a successful build")
    args = parser.parse_args(argv)
    if args.update_engine and (not args.engine_ref or args.check):
        parser.error("--update-engine requires --engine-ref and cannot be combined with --check")
    manifest = HERE / "engine.json"
    pin = json.loads(manifest.read_text(encoding="utf-8"))
    revision = resolve_revision(args.engine_root, args.engine_ref or pin["revision"])
    selected = dict(pin, revision=revision)
    result = generate(args.engine_root, selected)
    target = HERE.parents[1] / "src" / "stdRegex.cls"
    if args.check:
        if target.read_text(encoding="utf-8-sig") != result:
            raise SystemExit("stdRegex.cls needs regeneration; run Tools/stdRegex-build/build.py")
        print("stdRegex.cls matches the wrapper and pinned engine.")
    else:
        target.write_bytes(result.replace("\n", "\r\n").encode("utf-8"))
        if args.update_engine:
            manifest.write_text(json.dumps(selected, indent=2) + "\n", encoding="utf-8", newline="\n")
        print(f"Generated {target} from {revision}")
        if args.engine_ref and not args.update_engine:
            print("engine.json is unchanged. Use --update-engine to adopt this revision for future builds.")


if __name__ == "__main__":
    main()
