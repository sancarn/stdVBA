"""Summarize verified per-call batch timings; no third-party packages required."""
import csv
import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path

directory = Path(sys.argv[1])
files = [directory / 'samples.csv', directory / 'repeat' / 'samples.csv']
groups = defaultdict(list)
for run, path in enumerate(files, 1):
    with path.open(encoding='utf-8-sig', newline='') as stream:
        for row in csv.DictReader(stream):
            row['run'] = run
            groups[(row['scenario'], row['operation'], row['engine'])].append(row)

rows = []
for (scenario, operation, engine), measurements in groups.items():
    statuses = {m['status'] for m in measurements}
    status = next(iter(statuses)) if len(statuses) == 1 else 'inconsistent'
    result = dict(scenario=scenario, operation=operation, engine=engine,
                  status=status, input_chars=int(measurements[0]['input_chars']),
                  expected_matches=int(measurements[0]['expected_matches']),
                  runs=len(measurements), notes=sorted({m['note'] for m in measurements if m['note']}))
    if status == 'ok':
        samples = [float(m[f'sample_{i}_us']) for m in measurements for i in range(1, 6)]
        result.update(median_us=statistics.median(samples), min_us=min(samples),
                      max_us=max(samples), samples_us=samples,
                      iterations_by_run=[int(m['iterations']) for m in measurements],
                      run_medians_us=[statistics.median(float(m[f'sample_{i}_us']) for i in range(1, 6)) for m in measurements])
    rows.append(result)

metadata = json.loads((directory / 'repeat' / 'samples.json').read_text(encoding='utf-8-sig'))
summary = dict(metadata=metadata, rows=rows,
               statistic='Median of per-call batch means across two runs (10 batches per cell); allocation diagnostic has 5 batches in the second run only.')
(directory / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n', encoding='utf-8')
for operation in ['Compile + first Test', 'Test', 'MatchAll', 'Replace', 'Allocation baseline']:
    print(operation)
    for r in rows:
        if r['operation'] == operation:
            number = f"{r['median_us'] / 1000:.6f} ms" if r['status'] == 'ok' else r['status']
            print(f"  {r['scenario']}: {r['engine']} {number}")
