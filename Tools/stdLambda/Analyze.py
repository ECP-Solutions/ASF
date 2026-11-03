"""Summarize the per-call batch timings of two runs; no third-party packages required.

    python Tools/stdLambda/Analyze.py <run directory>

Reads samples.csv and repeat/samples.csv, writes summary.json and summary.md (median of the ten
per-call batch means of each cell, and the stdLambda/ASF ratio).
"""
import csv
import json
import math
import statistics
import sys
from collections import defaultdict
from pathlib import Path

directory = Path(sys.argv[1])
files = [directory / 'samples.csv', directory / 'repeat' / 'samples.csv']
groups = defaultdict(list)
order = []
for path in files:
    with path.open(encoding='utf-8-sig', newline='') as stream:
        for row in csv.DictReader(stream):
            key = (row['workload'], row['operation'], row['engine'])
            if key not in groups:
                order.append(key)
            groups[key].append(row)

rows = []
for key in order:
    workload, operation, engine = key
    measurements = groups[key]
    statuses = {m['status'] for m in measurements}
    status = next(iter(statuses)) if len(statuses) == 1 else 'inconsistent'
    result = dict(workload=workload, operation=operation, engine=engine, status=status, runs=len(measurements),
                  notes=sorted({m['note'] for m in measurements if m['note']}))
    if status == 'ok':
        samples = [float(m[f'sample_{i}_us']) for m in measurements for i in range(1, 6)]
        result.update(median_us=statistics.median(samples), min_us=min(samples), max_us=max(samples),
                      samples_us=samples, iterations_by_run=[int(m['iterations']) for m in measurements])
    rows.append(result)

metadata = json.loads((directory / 'repeat' / 'samples.json').read_text(encoding='utf-8-sig'))
summary = dict(metadata=metadata, rows=rows,
               statistic='Median of per-call batch means across two runs (10 batches per cell).')
(directory / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n', encoding='utf-8')


def fmt(us):
    return f'{us / 1000:.3f} ms' if us >= 1000 else f'{us:.1f} µs'


cells = {(r['workload'], r['operation'], r['engine']): r for r in rows}
lines = []
for operation in ['Compile', 'Run']:
    lines += [f'### {operation}', '', '| Workload | stdLambda | ASF | ASF / stdLambda |', '|---|---|---|---|']
    logs = []
    for workload in dict.fromkeys(r['workload'] for r in rows):
        a, b = cells.get((workload, operation, 'stdLambda')), cells.get((workload, operation, 'ASF'))
        ta = fmt(a['median_us']) if a and a['status'] == 'ok' else (a['status'] if a else '-')
        tb = fmt(b['median_us']) if b and b['status'] == 'ok' else (b['status'] if b else '-')
        ratio = ''
        if a and b and a['status'] == b['status'] == 'ok':
            q = b['median_us'] / a['median_us']
            logs.append(math.log(q))
            ratio = f'{q:.2f}×'
        lines.append(f'| {workload} | {ta} | {tb} | {ratio} |')
    if logs:
        lines.append(f'| **Geometric mean** | | | **{math.exp(sum(logs) / len(logs)):.2f}×** |')
    lines.append('')
text = '\n'.join(lines)
(directory / 'summary.md').write_text(text, encoding='utf-8')
print(text)
