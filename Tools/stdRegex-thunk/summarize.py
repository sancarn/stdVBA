"""Summarize the measured samples and create a standalone comparison chart."""
import json
from pathlib import Path
import statistics

HERE=Path(__file__).resolve().parent
def write(path,text):
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))

data=json.loads((HERE/'results/results.json').read_text())
assert data['status']=='passed' and not data['dialogs']
summary=[]
for case in dict.fromkeys(r['case'] for r in data['samples']):
    for operation in ['Test','List','ListArr','MatchAllArr']:
        record=dict(case=case,operation=operation)
        for engine in ['VBA','native']:
            rows=[r for r in data['samples'] if r['case']==case and r['operation']==operation+' '+engine]
            values=[r['ms_per_call'] for r in rows]
            assert len(values)==7
            record[engine]=dict(median_ms=statistics.median(values),min_ms=min(values),max_ms=max(values))
        record['speedup']=record['VBA']['median_ms']/record['native']['median_ms']
        summary.append(record)
write(HERE/'results/summary.json',json.dumps(summary,indent=2)+'\n')
html='''<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{font:14px var(--font-sans,system-ui);color:var(--foreground);margin:0}
p{color:var(--muted-foreground);line-height:1.5}select{font:inherit;color:var(--foreground);background:var(--card);border:1px solid var(--border);padding:6px}
svg{width:100%;height:380px}text{fill:var(--foreground);font:12px var(--font-sans,system-ui)}
table{border-collapse:collapse;width:100%;font-size:12px}td,th{text-align:right;padding:7px 3px;border-bottom:1px solid var(--border)}td:first-child,th:first-child{text-align:left}
</style></head><body>
<label>Measured operation <select id="op"><option>Test</option><option>List</option><option>ListArr</option><option>MatchAllArr</option></select></label>
<p>Native matcher speedup over current VBA (ratio of median call times; larger is faster).</p>
<svg id="chart" viewBox="0 0 700 380" role="img" aria-label="Native matcher speedup by workload"></svg>
<table><thead><tr><th>Workload</th><th>VBA ms</th><th>Native ms</th><th>Speedup</th></tr></thead><tbody id="rows"></tbody></table>
<p>Source: local Word 16.0, Windows x64, 1 November 2026. Seven rotated samples per operation; warmed patterns; batches calibrated to at least 50 ms. Returned values are allocated and released inside timing. Test stops at the first match. Compiler and public formatting remain VBA. These are workload-specific results, not a universal speed guarantee.</p>
<script>const data=DATA;
function render(){const rows=data.filter(x=>x.operation===document.getElementById('op').value),max=Math.ceil(Math.max(...rows.map(x=>x.speedup))/10)*10;
let svg='';for(let i=0;i<=4;i++){let x=240+i*400/4;svg+=`<line x1="${x}" x2="${x}" y1="15" y2="330" stroke="var(--border)"/><text x="${x}" y="350" text-anchor="middle">${(max*i/4).toFixed(0)}x</text>`}
rows.forEach((r,i)=>{let y=25+i*50;svg+=`<text x="0" y="${y+17}">${r.case.replace('characters','chars').replace(' named captures',' captures')}</text><rect x="240" y="${y}" width="${400*r.speedup/max}" height="25" fill="var(--chart-1)"/><text x="${245+400*r.speedup/max}" y="${y+18}">${r.speedup.toFixed(1)}x</text>`});
svg+='<text x="440" y="375" text-anchor="middle">Speedup (VBA time / native time)</text>';document.getElementById('chart').innerHTML=svg;
document.getElementById('rows').innerHTML=rows.map(r=>`<tr><td>${r.case}</td><td>${r.VBA.median_ms.toFixed(4)}</td><td>${r.native.median_ms.toFixed(4)}</td><td>${r.speedup.toFixed(2)}x</td></tr>`).join('');}
document.getElementById('op').onchange=render;render();</script></body></html>'''.replace('DATA',json.dumps(summary))
write(HERE/'results/comparison.html',html)
for r in summary:
    print(f"{r['case']} | {r['operation']}: {r['speedup']:.2f}x")
