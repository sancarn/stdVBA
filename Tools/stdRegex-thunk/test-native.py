from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).resolve().parents[2] / "tests/harnesses/stdRegex-thunk/test-native.py"), run_name="__main__")
