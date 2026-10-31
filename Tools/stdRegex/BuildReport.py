"""Create a self-contained HTML report from a benchmark summary (see Analyze.py).

    python Tools/stdRegex/BuildReport.py <run directory>

Adapted from stdVBA's Tools/stdRegex (branch codex/issue-141-regex-analysis). The tables,
chart and findings are computed from summary.json, so the report works for any engine set and
any environment (Excel or LibreOffice).
"""
import html
import json
import math
import sys
from datetime import datetime
from pathlib import Path

directory = Path(sys.argv[1])
data = json.loads((directory / 'summary.json').read_text(encoding='utf-8-sig'))
meta = data['metadata']
ORDER = ['Old stdRegex', 'Sihlfall', 'ASF 290ad3e', 'ASF v3.2.0', 'ASF']
present = {r['engine'] for r in data['rows'] if r['operation'] != 'Allocation baseline'}
engines = [e for e in ORDER if e in present] + sorted(present - set(ORDER))
not_run = [e for e in ORDER if e not in present]
scenarios = ['Literal scan', 'Short validation', 'Capture extraction', 'Alternation and ignore case', 'Multiline anchors', 'No-match scan', 'Word boundaries', 'Backreference']
operations = ['Test', 'MatchAll', 'Replace', 'Compile + first Test']
definitions = [
    ('needle', '2,054 characters; match after a 1,024-character prefix', 'g'),
    (r'^[A-Z]{2}\d{6}$', 'GB123456; 8 characters', 'g'),
    (r'([A-Z]{3})-(\d{4})', '100 ABC-1234 records; 1,000 characters', 'g'),
    ('(ERROR|WARN|INFO): ([A-Za-z]+)', '100 mixed-case messages; 1,300 characters', 'gi'),
    (r'^([A-Z]+):(\d+)$', '100 ITEM:123 lines; 900 characters', 'gm'),
    (r'Z\d{4}', '1,024 a characters; zero matches', 'g'),
    (r'\bcat\b', 'a cat scatter cat!; two matches', 'g'),
    (r'^(ab)\1$', 'abab; one match', 'g'),
]
descriptions = {
    'Test': 'Reused regex; first-match Boolean checked against the expected result.',
    'MatchAll': 'All matches and numbered captures checked. stdRegex returns dictionaries; ASF returns collections, so the output costs differ.',
    'Replace': 'Reused regex; global replacement. The complete output string, including capture substitutions, was checked.',
    'Compile + first Test': 'Fresh object, initialization and first Test together after global warm-up; constructor caches are bypassed.',
}
SOURCES = {'Old stdRegex': 'legacy_stdregex_commit', 'Sihlfall': 'sihlfall_stdregex_commit', 'ASF 290ad3e': 'asf_290ad3e_commit',
           'ASF v3.2.0': 'asf_v320_commit', 'ASF': 'asf_commit'}


def esc(value):
    return html.escape(str(value), quote=True)


def find(scenario, operation, engine):
    return next((r for r in data['rows'] if (r['scenario'], r['operation'], r['engine']) == (scenario, operation, engine)), None)


def ms(value):
    return f'{value:,.3f}' if value >= 0.01 else f'{value:.4f}'


def speedups(fast, slow, operation):
    """Ratios slow/fast over the workloads where both produced the expected output."""
    out = []
    for scenario in scenarios:
        a, b = find(scenario, operation, fast), find(scenario, operation, slow)
        if a and b and a['status'] == 'ok' and b['status'] == 'ok':
            out.append((scenario, b['median_us'] / a['median_us']))
    return out


def geomean(values):
    return math.exp(sum(math.log(v) for v in values) / len(values))


def pair_finding(fast, slow):
    parts = []
    for operation in operations:
        ratios = speedups(fast, slow, operation)
        if not ratios:
            continue
        values = [r for _, r in ratios]
        best = max(ratios, key=lambda x: x[1])
        parts.append(f'{esc(operation)}: {geomean(values):.1f}× on average ({min(values):.1f}–{max(values):.1f}×, best on {esc(best[0].lower())})')
    if not parts:
        return ''
    return f'<p><strong>{esc(fast)} against {esc(slow)}:</strong> ' + '; '.join(parts) + '. Geometric means over the workloads where both give the expected output.</p>'


findings = []
for fast, slow in [('ASF', 'ASF 290ad3e'), ('ASF', 'ASF v3.2.0'), ('ASF', 'Sihlfall'), ('ASF', 'Old stdRegex')]:
    if fast in engines and slow in engines:
        findings.append(pair_finding(fast, slow))
mismatches = {}
for r in data['rows']:
    if r['status'] in ('mismatch', 'inconsistent'):
        mismatches.setdefault(r['engine'], set()).add(r['scenario'])
for engine in engines:
    if engine in mismatches:
        findings.append(f'<p><strong>Compatibility:</strong> {esc(engine)} fails the {esc(" and ".join(s.lower() for s in scenarios if s in mismatches[engine]))} probes; its cells there are excluded from the comparisons.</p>')
if 'ASF' in engines and 'ASF' not in mismatches:
    findings.append('<p><strong>ASF</strong> gives the expected output in every cell, word boundaries and backreferences included.</p>')

sections = []
for operation in operations:
    rows = []
    for scenario in scenarios:
        cells = ['<th scope="row">' + esc(scenario) + '</th>']
        for engine in engines:
            result = find(scenario, operation, engine)
            if result is None:
                cells.append('<td class="mismatch">Not run</td>')
            elif result['status'] == 'ok':
                detail = f"Batch-mean range: {result['min_us']/1000:.6f}–{result['max_us']/1000:.6f} ms. Run medians: " + ', '.join(f'{v/1000:.6f} ms' for v in result['run_medians_us'])
                cells.append(f'<td title="{esc(detail)}">{ms(result["median_us"]/1000)}</td>')
            elif result['status'] == 'mismatch':
                cells.append('<td class="mismatch" title="' + esc('; '.join(result['notes'])) + '">Not equivalent</td>')
            else:
                cells.append('<td class="mismatch" title="' + esc('; '.join(result['notes'])) + '">' + esc(result['status'].capitalize()) + '</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')
    sections.append('<section><h2>' + esc(operation) + '<span class="unit">milliseconds per call</span></h2><p class="muted">' + descriptions[operation]
                    + '</p><div class="table-wrap"><table><thead><tr><th>Workload</th>' + ''.join('<th>' + esc(e) + ' (ms)</th>' for e in engines)
                    + '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div></section>')

pattern_rows = ''.join('<tr><th scope="row">' + esc(name) + '</th><td class="pattern"><code>' + esc(pattern) + '</code></td><td>' + esc(description) + '</td><td><code>' + flags + '</code></td></tr>'
                       for name, (pattern, description, flags) in zip(scenarios, definitions))
allocations = [r for r in data['rows'] if r['operation'] == 'Allocation baseline' and r['status'] == 'ok']
allocation_text = '; '.join(f"{esc(r['engine'].lower())}: {r['median_us']/1000:.3f} ms" for r in allocations)
compact = [{'scenario': r['scenario'], 'operation': r['operation'], 'engine': r['engine'], 'status': r['status'], 'input_chars': r['input_chars'],
            **({'ms': r['median_us'] / 1000} if r['status'] == 'ok' else {})} for r in data['rows'] if r['operation'] != 'Allocation baseline']
measured = datetime.fromisoformat(meta['measured_at_utc'].replace('Z', '+00:00'))
date_text = f'{measured.day} {measured:%B %Y}'
environment = f"{meta.get('office', 'Excel ' + str(meta.get('excel_version', '')))} · {meta.get('operating_system', '')} · {meta.get('cpu', '')}"
libreoffice = 'LibreOffice' in str(meta.get('office', ''))
versions = ', '.join(f'{esc(e)} <code>{esc(str(meta[SOURCES[e]])[:7])}</code>' for e in engines if SOURCES.get(e) in meta)
env_note = ('<p>This run used the LibreOffice edition of ASF, generated from the same VBA sources (see <code>libreoffice/</code>), '
            'so absolute times are LibreOffice Basic times, several times slower than VBA; compare engines within the run. '
            + (f'{esc(", ".join(not_run))} need Windows (VBScript, <code>Scripting.Dictionary</code>) and run in Excel only, with <code>Benchmark.ps1</code>. ' if not_run else '')
            + 'The clock is LibreOffice\'s millisecond tick count, so a 30 ms batch has a resolution of about 3%.</p>') if libreoffice else ''

document = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Regex engine benchmarks — __DATE__</title>
<style>
:root{color-scheme:light dark;--bg:#fafafa;--surface:#fff;--fg:#222;--muted:#626262;--line:#dcdcdc;--accent:#2366ad;--tint:#edf3fa}
@media(prefers-color-scheme:dark){:root{--bg:#191b1e;--surface:#202327;--fg:#edf0f3;--muted:#b0b6be;--line:#3c424a;--accent:#8bbcff;--tint:#252f3c}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif}main{max-width:1120px;margin:0 auto;padding:44px 28px 64px}h1{font-size:30px;line-height:1.2;margin:8px 0 16px;letter-spacing:-.025em}h2{font-size:20px;line-height:1.35;margin:0 0 10px}p{margin:8px 0 14px}.eyebrow,.muted,.unit{color:var(--muted)}.eyebrow{font-size:13px}.unit{display:block;font-size:13px;font-weight:400;margin-top:3px}section{margin-top:36px}.findings{border-left:3px solid var(--accent);padding:2px 0 2px 20px;margin-top:26px}.findings p{margin:9px 0}.chart-frame{border:1px solid var(--line);background:var(--surface);padding:20px 16px 10px;margin-top:15px}.chart-scroll{overflow-x:auto}svg{display:block;width:100%;min-width:640px}svg text{font-family:inherit;fill:var(--fg);font-size:12px}svg .secondary{fill:var(--muted)}svg .grid{stroke:var(--line);stroke-width:1}.legend{display:flex;flex-wrap:wrap;gap:20px;font-size:13px}.legend span{display:flex;align-items:center;gap:7px}.dot{width:10px;height:10px;border-radius:50%;display:inline-block}label{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:16px}select{font:inherit;color:var(--fg);background:var(--surface);border:1px solid var(--line);border-radius:4px;padding:6px 10px}.table-wrap{overflow-x:auto;border:1px solid var(--line);background:var(--surface)}table{width:100%;border-collapse:collapse;font-size:14px}th,td{padding:11px 15px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}thead th{font-size:12px;color:var(--muted);font-weight:600}th:first-child{text-align:left}tbody th{font-weight:500}tr:last-child td,tr:last-child th{border-bottom:0}.mismatch{font-size:12px;color:var(--muted)}code{font:13px/1.5 ui-monospace,Consolas,monospace}.patterns td{text-align:left;white-space:normal}.patterns .pattern{white-space:nowrap}.caption{font-size:12px;color:var(--muted);margin:12px 0 4px}details{border-top:1px solid var(--line);margin-top:36px;padding-top:20px}summary{font-weight:600;cursor:pointer}.method{max-width:950px}.method p{margin-top:14px}a{color:var(--accent)}footer{margin-top:32px;padding-top:18px;border-top:1px solid var(--line);color:var(--muted);font-size:12px}
@media(max-width:650px){main{padding:25px 16px 40px}h1{font-size:26px}th,td{padding:10px 11px}.chart-frame{padding:14px 10px}.findings{padding-left:14px}}
@media print{body{background:#fff;color:#111}main{max-width:none;padding:0}section{break-inside:avoid}select,label{display:none}svg{min-width:0}.table-wrap{overflow:visible}}
</style></head><body><main>
<header><div class="eyebrow">ASF · Tools/stdRegex · measured __DATE__</div><h1>Regex engine benchmarks</h1>
<p class="muted">__ENV__. Eight patterns, four operations, __NENGINES__ implementations, two independent runs.</p></header>
<div class="findings">__FINDINGS__</div>
<section aria-labelledby="chart-title"><h2 id="chart-title">Median time per call</h2><p class="muted">Workloads where every implementation gives the expected output. Lower is faster; the time axis is logarithmic.</p>
<label for="operation">Operation <select id="operation"><option>Test</option><option>MatchAll</option><option>Replace</option><option>Compile + first Test</option></select></label>
<div class="chart-frame"><div class="legend" id="legend"></div>
<div class="chart-scroll"><svg id="chart" viewBox="0 0 940 375" role="img" aria-label="Median time per call"></svg></div>
<p class="caption">Median of 10 per-call batch means across two runs. Hover a point for its value. Exact values, batch ranges and separate run medians are in the tables below.</p></div></section>
__SECTIONS__
<section><h2>Patterns and input sizes</h2><div class="table-wrap"><table class="patterns"><thead><tr><th>Workload</th><th>Pattern</th><th>Input</th><th>Flags</th></tr></thead><tbody>__PATTERNS__</tbody></table></div></section>
<details open><summary>Method, interpretation and limits</summary><div class="method">
<p>The workloads, checks and timing procedure are those of stdVBA's regex benchmark (<code>Tools/stdRegex</code>, branch <code>codex/issue-141-regex-analysis</code>), extended with more ASF versions. Every timing sample was accepted only after output checks: the expected Boolean for Test, every numbered capture and the match count for MatchAll, the complete replacement string for Replace. Cells marked "Not equivalent" failed that check and are left out of the comparisons.</p>
<p>The clock runs inside Basic (__TIMER__), excluding automation calls. Each cell is calibrated to at least 30 ms per batch, up to 16,384 iterations, followed by five timed batches. The whole experiment ran twice in separate processes; the tables show medians of the resulting ten batch averages per call, not individual-call latency percentiles. Engine order rotates between workload/operation cells.</p>
<p>Compile + first Test builds a fresh object, so constructor caches are bypassed; ASF's compiled-pattern cache (v3.2.0 and later) belongs to each object unless shared through <code>SharedCache</code>, so every fresh object compiles. Other operations reuse preinitialized objects. Native processing limits are raised to 2,147,483,647. These small inputs do not test pathological backtracking.</p>
<p><strong>Capture output is a confounder</strong> for MatchAll: stdRegex returns dictionaries, ASF returns collections. Building 100 result containers without any regex work took __ALLOC__ in this run.</p>
__ENVNOTE__
<p>Source versions: __VERSIONS__.</p>
</div></details>
<footer>Self-contained report: all measurements, charts, styles and scripts are embedded. Table units: milliseconds per call. Generated by Tools/stdRegex/BuildReport.py from summary.json.</footer>
</main><script id="benchmark-data" type="application/json">__DATA__</script><script>
'use strict';
const data=JSON.parse(document.getElementById('benchmark-data').textContent);
const engines=__ENGINES__;
const palette=['var(--fg)','var(--accent)','#c0392b','#8e44ad','#16a085'];
const svg=document.getElementById('chart');
const picker=document.getElementById('operation');
const ns='http://www.w3.org/2000/svg';
document.getElementById('legend').innerHTML=engines.map((e,i)=>'<span><i class="dot" style="background:'+palette[i%palette.length]+'"></i>'+e.replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</span>').join('');
function node(tag,attrs,text,parent=svg){const element=document.createElementNS(ns,tag);for(const [key,value] of Object.entries(attrs||{}))element.setAttribute(key,String(value));if(text!==undefined)element.textContent=text;parent.appendChild(element);return element;}
function draw(){
  const operation=picker.value;
  const all=data.filter(r=>r.operation===operation);
  const scenarios=[...new Set(all.map(r=>r.scenario))].filter(s=>engines.every(e=>all.some(r=>r.scenario===s&&r.engine===e&&r.status==='ok')));
  const rows=all.filter(r=>r.status==='ok'&&scenarios.includes(r.scenario));
  if(!rows.length){svg.replaceChildren();return;}
  const values=rows.map(r=>r.ms);
  const low=Math.floor(Math.log10(Math.min(...values))),high=Math.max(low+1,Math.ceil(Math.log10(Math.max(...values))));
  const x=value=>230+(Math.log10(value)-low)/(high-low)*650;
  const height=Math.max(375,90+scenarios.length*47);
  svg.setAttribute('viewBox','0 0 940 '+height);
  svg.replaceChildren();svg.setAttribute('aria-label','Median '+operation+' time in milliseconds');
  node('text',{x:8,y:19,class:'secondary'},'Workload (input characters)');
  for(let exponent=low;exponent<=high;exponent++){
    const tick=10**exponent;
    node('line',{x1:x(tick),x2:x(tick),y1:30,y2:height-65,class:'grid'});
    node('text',{x:x(tick),y:height-42,'text-anchor':'middle',class:'secondary'},tick<1?tick.toFixed(-exponent):String(tick));
  }
  scenarios.forEach((scenario,index)=>{
    const y=50+index*47,first=rows.find(r=>r.scenario===scenario);
    node('text',{x:8,y:y-2},scenario);node('text',{x:8,y:y+14,class:'secondary'},first.input_chars.toLocaleString('en-GB')+' chars');
    node('line',{x1:230,x2:880,y1:y,y2:y,class:'grid'});
    engines.forEach((engine,series)=>{
      const result=rows.find(r=>r.scenario===scenario&&r.engine===engine);
      if(!result)return;
      const point=node('circle',{cx:x(result.ms),cy:y,r:5,fill:palette[series%palette.length],stroke:'var(--surface)','stroke-width':1});
      node('title',{},engine+': '+result.ms.toLocaleString('en-GB',{maximumFractionDigits:6})+' ms/call',point);
    });
  });
  node('text',{x:555,y:height-9,'text-anchor':'middle',class:'secondary'},'Time per call (ms, logarithmic scale)');
}
picker.addEventListener('change',draw);draw();
</script></body></html>'''
replacements = {
    '__DATE__': esc(date_text), '__ENV__': esc(environment), '__NENGINES__': str(len(engines)),
    '__FINDINGS__': '\n'.join(findings), '__SECTIONS__': '\n'.join(sections), '__PATTERNS__': pattern_rows,
    '__TIMER__': esc(meta.get('timer', '')), '__ALLOC__': allocation_text or 'n/a', '__ENVNOTE__': env_note,
    '__VERSIONS__': versions or 'see summary.json',
    '__ENGINES__': json.dumps(engines),
    '__DATA__': json.dumps(compact, separators=(',', ':')).replace('<', '\\u003c'),
}
for key, value in replacements.items():
    document = document.replace(key, value)
output = directory / 'report.html'
output.write_text(document, encoding='utf-8', newline='\n')
print(output.resolve())
