"""Create a self-contained HTML report from the recorded benchmark summary."""
import html
import json
import sys
from pathlib import Path

directory = Path(sys.argv[1])
data = json.loads((directory / 'summary.json').read_text(encoding='utf-8-sig'))
engines = ['Old stdRegex', 'Sihlfall', 'ASF']
scenarios = ['Literal scan', 'Short validation', 'Capture extraction', 'Alternation and ignore case', 'Multiline anchors', 'No-match scan', 'Word boundaries', 'Backreference']
operations = ['Test', 'MatchAll', 'Replace', 'Compile + first Test']
definitions = [
    ('needle', '2,054 characters; match after a 1,024-character prefix', 'g'),
    (r'^[A-Z]{2}\d{6}$', 'GB123456; 8 characters', 'g'),
    (r'([A-Z]{3})-(\d{4})', '100 ABC-1234 records; 1,000 characters', 'g'),
    ('(ERROR|WARN|INFO): ([A-Za-z]+)', '100 mixed-case messages; 1,250 characters', 'gi'),
    (r'^([A-Z]+):(\d+)$', '100 ITEM:123 lines; 900 characters', 'gm'),
    (r'Z\d{4}', '1,024 a characters; zero matches', 'g'),
    (r'\bcat\b', 'a cat scatter cat!; two matches', 'g'),
    (r'^(ab)\1$', 'abab; one match', 'g'),
]
def esc(value):
    return html.escape(str(value), quote=True)

def find(scenario, operation, engine):
    return next(r for r in data['rows'] if (r['scenario'], r['operation'], r['engine']) == (scenario, operation, engine))

sections = []
descriptions = {
    'Test': 'Reused regex; first-match Boolean checked against the expected result.',
    'MatchAll': 'All matches and numbered captures checked. stdRegex returns dictionaries; ASF returns collections. These output costs differ.',
    'Replace': 'Reused regex; global replacement. The complete output string, including capture substitutions, was checked.',
    'Compile + first Test': 'Fresh object, initialization, and first Test together after global warm-up. Includes the first search because VBScript compiles lazily; constructor caches are bypassed.',
}
for operation in operations:
    rows = []
    for scenario in scenarios:
        cells = ['<th scope="row">' + esc(scenario) + '</th>']
        for engine in engines:
            result = find(scenario, operation, engine)
            if result['status'] == 'ok':
                detail = f"Batch-mean range: {result['min_us']/1000:.6f}–{result['max_us']/1000:.6f} ms. Run medians: " + ', '.join(f'{v/1000:.6f} ms' for v in result['run_medians_us'])
                cells.append(f'<td title="{esc(detail)}">{result["median_us"]/1000:,.3f}</td>')
            else:
                cells.append('<td class="mismatch" title="' + esc('; '.join(result['notes'])) + '">Not equivalent</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')
    sections.append('<section><h2>' + esc(operation) + '<span class="unit">milliseconds per call</span></h2><p class="muted">' + descriptions[operation] + '</p><div class="table-wrap"><table><thead><tr><th>Workload</th>' + ''.join('<th>' + esc(e) + ' (ms)</th>' for e in engines) + '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div></section>')

pattern_rows = ''.join('<tr><th scope="row">' + esc(name) + '</th><td class="pattern"><code>' + esc(pattern) + '</code></td><td>' + esc(description) + '</td><td><code>' + flags + '</code></td></tr>' for name, (pattern, description, flags) in zip(scenarios, definitions))
compact = [{'scenario': r['scenario'], 'operation': r['operation'], 'engine': r['engine'], 'status': r['status'], 'input_chars': r['input_chars'], **({'ms': r['median_us'] / 1000} if r['status'] == 'ok' else {})} for r in data['rows']]
document = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>stdRegex engine benchmarks — 27 October 2026</title>
<style>
:root{color-scheme:light dark;--bg:#fafafa;--surface:#fff;--fg:#222;--muted:#626262;--line:#dcdcdc;--accent:#2366ad;--tint:#edf3fa}
@media(prefers-color-scheme:dark){:root{--bg:#191b1e;--surface:#202327;--fg:#edf0f3;--muted:#b0b6be;--line:#3c424a;--accent:#8bbcff;--tint:#252f3c}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif}main{max-width:1120px;margin:0 auto;padding:44px 28px 64px}h1{font-size:30px;line-height:1.2;margin:8px 0 16px;letter-spacing:-.025em}h2{font-size:20px;line-height:1.35;margin:0 0 10px}h3{font-size:16px;margin:0 0 8px}p{margin:8px 0 14px}.eyebrow,.muted,.unit{color:var(--muted)}.eyebrow{font-size:13px}.lead{font-size:18px;max-width:850px}.unit{display:block;font-size:13px;font-weight:400;margin-top:3px}section{margin-top:36px}.findings{border-left:3px solid var(--accent);padding:2px 0 2px 20px;margin-top:26px}.findings p{margin:9px 0}.chart-frame{border:1px solid var(--line);background:var(--surface);padding:20px 16px 10px;margin-top:15px}.chart-scroll{overflow-x:auto}svg{display:block;width:100%;min-width:640px}svg text{font-family:inherit;fill:var(--fg);font-size:12px}svg .secondary{fill:var(--muted)}svg .grid{stroke:var(--line);stroke-width:1}.legend{display:flex;flex-wrap:wrap;gap:20px;font-size:13px}.legend span{display:flex;align-items:center;gap:7px}.dot{width:10px;height:10px;border-radius:50%;background:var(--fg);display:inline-block}.dot.sihlfall{background:var(--accent)}.dot.asf{background:var(--surface);border:2px solid var(--fg)}label{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:16px}select{font:inherit;color:var(--fg);background:var(--surface);border:1px solid var(--line);border-radius:4px;padding:6px 10px}.table-wrap{overflow-x:auto;border:1px solid var(--line);background:var(--surface)}table{width:100%;border-collapse:collapse;font-size:14px}th,td{padding:11px 15px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}thead th{font-size:12px;color:var(--muted);font-weight:600}th:first-child{text-align:left}tbody th{font-weight:500}tr:last-child td,tr:last-child th{border-bottom:0}.mismatch{font-size:12px;color:var(--muted)}code{font:13px/1.5 ui-monospace,Consolas,monospace}.patterns td{text-align:left;white-space:normal}.patterns .pattern{white-space:nowrap}.caption{font-size:12px;color:var(--muted);margin:12px 0 4px}details{border-top:1px solid var(--line);margin-top:36px;padding-top:20px}summary{font-weight:600;cursor:pointer}.method{max-width:950px}.method p{margin-top:14px}a{color:var(--accent)}footer{margin-top:32px;padding-top:18px;border-top:1px solid var(--line);color:var(--muted);font-size:12px}
@media(max-width:650px){main{padding:25px 16px 40px}h1{font-size:26px}.lead{font-size:16px}th,td{padding:10px 11px}.chart-frame{padding:14px 10px}.findings{padding-left:14px}}
@media print{body{background:#fff;color:#111}main{max-width:none;padding:0}section{break-inside:avoid}details .method{display:block}select,label{display:none}.chart-frame{border-color:#ccc}svg{min-width:0}.table-wrap{overflow:visible}footer{margin-top:20px}}
</style></head><body><main>
<header><div class="eyebrow">stdVBA · measured 27 October 2026</div><h1>Regex performance trade-offs</h1>
<p class="lead">Sihlfall beats ASF for <code>Test</code> and replacement. The old COM engine is faster on reused matching. ASF's faster capture-returning API has a different output cost.</p>
<p class="muted">Excel 16.0 build 20618 · x64 · Intel Core i7-9700KF. Eight patterns, four operations, three implementations, two independent workbook runs.</p></header>
<div class="findings"><p><strong>100 captured records:</strong> MatchAll takes 175 ms with Sihlfall, 177 ms with old stdRegex, and 31 ms with ASF. Both stdRegex versions create dictionaries; ASF returns lighter collections.</p>
<p><strong>Replace those records:</strong> Sihlfall takes 1.98 ms, old stdRegex 0.025 ms, and ASF 37.4 ms.</p>
<p><strong>Compatibility:</strong> ASF fails the word-boundary and numeric-backreference probes. Incorrect outputs are excluded from timing comparisons.</p></div>
<section aria-labelledby="chart-title"><h2 id="chart-title">Median time per call</h2><p class="muted">Six workloads with equivalent results across all three implementations. Lower is faster; the time axis is logarithmic.</p>
<label for="operation">Operation <select id="operation"><option>Test</option><option>MatchAll</option><option>Replace</option><option>Compile + first Test</option></select></label>
<div class="chart-frame"><div class="legend"><span><i class="dot"></i>Old stdRegex</span><span><i class="dot sihlfall"></i>Sihlfall</span><span><i class="dot asf"></i>ASF</span></div>
<div class="chart-scroll"><svg id="chart" viewBox="0 0 940 375" role="img" aria-label="Median Test time in milliseconds for six common workloads"></svg></div>
<p class="caption">Median of 10 per-call batch means across two runs. Hover a point for its value. Exact values, batch ranges, and separate run medians are available in the tables below.</p></div></section>
__SECTIONS__
<section><h2>Patterns and input sizes</h2><div class="table-wrap"><table class="patterns"><thead><tr><th>Workload</th><th>Pattern</th><th>Input</th><th>Flags</th></tr></thead><tbody>__PATTERNS__</tbody></table></div></section>
<details open><summary>Method, interpretation, and limits</summary><div class="method">
<p>All timing samples were accepted only after output checks: the expected Boolean for Test, every numbered capture and match count for MatchAll, or the complete replacement string. Both word-boundary and backreference probes fail on ASF, producing eight excluded cells per run. The exact known Office-crashing issue 141 pattern was not executed in the old engine; the safer <code>^(ab)\1$</code> probe was used.</p>
<p>QueryPerformanceCounter measures inside VBA, excluding Excel automation calls. Each cell is warmed and calibrated to at least 30 ms per batch, up to 16,384 iterations, followed by five timed batches. The complete experiment was repeated in a separate workbook run; the tables show medians of the resulting ten batch averages per call, not individual-call latency percentiles. Engine order rotates between workload/operation cells.</p>
<p>Fresh construction plus first Test includes the first search because VBScript compiles lazily; constructor cache hits are excluded. Other operations reuse preinitialized objects. Shared static initialization is warmed before timing. Normal wrapper work and intermediate result disposal are included.</p>
<p><strong>Capture output is an important confounder.</strong> The migrated and old stdRegex APIs create dictionaries, whereas ASF returns collections. In a separate five-batch diagnostic, constructing 100 comparable containers without any regex work took 67.2 ms for dictionaries and 0.311 ms for collections. The diagnostic discards each container before creating the next, whereas real MatchAll retains all results, so it is not an exact decomposition of MatchAll cost or proof that ASF's matcher is faster.</p>
<p>The migration is a compatibility/reliability choice, not a performance upgrade over VBScript. Short validation Test calls took about 17 µs for Sihlfall versus 4.7 µs for old stdRegex and 165 µs for ASF. Unanchored 1,024-character scans took Sihlfall about 1.1 ms versus the old engine's 0.005 ms and ASF's 25–26 ms. Match-result construction is a separate optimization target.</p>
<p>Native processing limits were raised to 2,147,483,647 to prevent accumulated ASF replacement budgets from distorting repeated runs. The old engine has no matching budget. These small representative inputs do not test pathological backtracking, and the results are specific to this computer. Background load and COM allocation vary; table tooltips show batch ranges and individual run medians.</p>
<p>Source versions: migrated stdRegex <code>04e7d04</code>, sihlfall/vba-regex <code>1619cd7</code>, old stdRegex <code>abf14dd</code>, ASF <code>290ad3e</code>. The existing 78 compatibility assertions passed before each benchmark. Temporary Excel access and debugger settings were restored; test processes were closed.</p>
</div></details>
<footer>Self-contained report: all measurements, charts, styles, and scripts are embedded. No internet connection or additional files required. Table units: milliseconds per call; three decimal places. Source: local benchmark recordings, 27 October 2026.</footer>
</main><script id="benchmark-data" type="application/json">__DATA__</script><script>
'use strict';
const data=JSON.parse(document.getElementById('benchmark-data').textContent);
const scenarios=['Literal scan','Short validation','Capture extraction','Alternation and ignore case','Multiline anchors','No-match scan'];
const engines=['Old stdRegex','Sihlfall','ASF'];
const svg=document.getElementById('chart');
const picker=document.getElementById('operation');
const ns='http://www.w3.org/2000/svg';
function node(tag,attrs,text,parent=svg){const element=document.createElementNS(ns,tag);for(const [key,value] of Object.entries(attrs||{}))element.setAttribute(key,String(value));if(text!==undefined)element.textContent=text;parent.appendChild(element);return element;}
function draw(){
  const operation=picker.value;
  const rows=data.filter(r=>r.operation===operation&&r.status==='ok'&&scenarios.includes(r.scenario));
  const values=rows.map(r=>r.ms);
  const low=Math.floor(Math.log10(Math.min(...values))),high=Math.ceil(Math.log10(Math.max(...values)));
  const x=value=>230+(Math.log10(value)-low)/(high-low)*650;
  svg.replaceChildren();svg.setAttribute('aria-label','Median '+operation+' time in milliseconds for six common workloads');
  node('text',{x:8,y:19,class:'secondary'},'Workload (input characters)');
  for(let exponent=low;exponent<=high;exponent++){
    const tick=10**exponent;
    node('line',{x1:x(tick),x2:x(tick),y1:30,y2:310,class:'grid'});
    node('text',{x:x(tick),y:333,'text-anchor':'middle',class:'secondary'},tick<1?tick.toFixed(-exponent):String(tick));
  }
  scenarios.forEach((scenario,index)=>{
    const y=50+index*47,first=rows.find(r=>r.scenario===scenario);
    node('text',{x:8,y:y-2},scenario);node('text',{x:8,y:y+14,class:'secondary'},first.input_chars.toLocaleString('en-GB')+' chars');
    node('line',{x1:230,x2:880,y1:y,y2:y,class:'grid'});
    engines.forEach((engine,series)=>{
      const result=rows.find(r=>r.scenario===scenario&&r.engine===engine);
      const point=node('circle',{cx:x(result.ms),cy:y,r:5,fill:series===1?'var(--accent)':series===0?'var(--fg)':'var(--surface)',stroke:series===1?'var(--accent)':'var(--fg)','stroke-width':1.5});
      node('title',{},engine+': '+result.ms.toLocaleString('en-GB',{maximumFractionDigits:6})+' ms/call',point);
    });
  });
  node('text',{x:555,y:366,'text-anchor':'middle',class:'secondary'},'Time per call (ms, logarithmic scale)');
}
picker.addEventListener('change',draw);draw();
</script></body></html>'''
document = document.replace('__SECTIONS__', '\n'.join(sections)).replace('__PATTERNS__', pattern_rows).replace('__DATA__', json.dumps(compact, separators=(',', ':')).replace('<', '\\u003c'))
output = directory / 'report.html'
output.write_text(document, encoding='utf-8', newline='\n')
print(output.resolve())
