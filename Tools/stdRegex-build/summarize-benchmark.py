from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).resolve().parents[2] / "tests/harnesses/stdRegex-build/summarize-benchmark.py"), run_name="__main__")
