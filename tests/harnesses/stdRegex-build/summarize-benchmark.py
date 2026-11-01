"""Summarize measured per-call samples; no synthetic timings."""
from collections import defaultdict
import json
from pathlib import Path
from statistics import median
import sys


def summarize(folder):
    samples = json.loads((folder / 'samples.json').read_text(encoding='utf-8-sig'))
    environment = json.loads((folder / 'environment.json').read_text(encoding='utf-8-sig'))
    groups = defaultdict(list)
    for row in samples:
        groups[(row['case'], row['operation'])].append(row['ms_per_call'])
    cases = []
    for case in dict.fromkeys(row['case'] for row in samples):
        timing = {}
        for (label, op), values in groups.items():
            if label == case:
                timing[op] = {'median_ms': median(values), 'min_ms': min(values), 'max_ms': max(values)}
        comparisons = []
        for before, after in [('List before', 'List after'), ('ListArr before', 'ListArr after'), ('MatchAll', 'MatchAllArr')]:
            old, new = timing[before]['median_ms'], timing[after]['median_ms']
            comparisons.append({'before': before, 'after': after, 'before_ms': old, 'after_ms': new,
                                'speedup': old / new, 'time_reduction_percent': 100 * (old - new) / old})
        cases.append({'case': case, 'timings': timing, 'comparisons': comparisons})
    result = {'environment': environment, 'cases': cases}
    (folder / 'summary.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    lines = ['# stdRegex before/after benchmark', '',
             f"Host: {environment['host_application']} {environment['office_version']}, {environment['vba_architecture']}.",
             f"Measured UTC: {environment['measured_at_utc']}.",
             f"Baseline: `{environment['baseline_commit']}` (pure VBA engine, before list optimizations).", '',
             f"{environment['samples']} samples per operation/case; medians in milliseconds per call. " + environment['methodology'], '',
             'ListArr uses two output columns. MatchAllArr materializes all numbered captures; MatchAll additionally creates dictionaries and named entries. Their contracts differ.', '']
    for case in cases:
        lines += ['## ' + case['case'], '']
        for item in case['comparisons']:
            change = item['time_reduction_percent']
            direction = 'less time' if change >= 0 else 'more time'
            lines.append(f"- {item['before']} → {item['after']}: **{item['before_ms']:.3f} → {item['after_ms']:.3f} ms**, {item['speedup']:.2f}× relative speed ({abs(change):.1f}% {direction}).")
        lines.append('')
    lines += ['These measurements describe this host and these workloads. Sparse/no-match workloads spend more time in the unchanged search engine, so wrapper-only improvements may be small or negative.', '']
    (folder / 'results.md').write_text('\n'.join(lines), encoding='utf-8')
    print('\n'.join(lines).replace('\u2192', '->').replace('\u00d7', 'x'))


if __name__ == '__main__':
    summarize(Path(sys.argv[1]))
