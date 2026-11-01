from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).resolve().parents[2] / "tests/native/test_window_thunk.py"), run_name="__main__")
