from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).resolve().parents[2] / "tests/tooling/test_regex_build.py"), run_name="__main__")
