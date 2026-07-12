#!/usr/bin/env python3
"""
Decorpot - Initiation BOM Generator (web app)
Upload an ERP quotation CSV, get the 3-sheet Initiation BOM (Lists/Price/BOM).
Master data + template are built in; team only uploads the CSV.
"""
import os
import io
import uuid
import tempfile
import threading
import time
from flask import Flask, request, send_file, render_template_string, jsonify

import initiation_bom as ib

app = Flask(__name__)

BASE = os.path.dirname(os.path.abspath(__file__))
MASTER = os.path.join(BASE, "master_data.xlsx")
TEMPLATE = os.path.join(BASE, "bom_template.xlsx")

# in-memory job store (stateless per file; cleared after download)
JOBS = {}

PAGE = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Initiation BOM Generator · Decorpot</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#f3efe7; --panel:#ffffff; --ink:#211f1b; --muted:#7c766c;
    --line:#e7e2d7; --brand:#e35b26; --brand-deep:#b8410f; --ok:#2f7d55;
    --shadow:0 18px 48px -24px rgba(50,40,25,.35);
  }
  *{box-sizing:border-box}
  html,body{height:100%}
  body{margin:0;font-family:"Inter",system-ui,sans-serif;color:var(--ink);
    background:
      radial-gradient(1100px 500px at 82% -8%, #fbe6d8 0%, rgba(251,230,216,0) 55%),
      radial-gradient(900px 500px at 0% 110%, #eef0e6 0%, rgba(238,240,230,0) 50%),
      var(--bg);
    display:flex;align-items:center;justify-content:center;padding:28px}
  .wrap{width:100%;max-width:620px}
  .brandrow{display:flex;align-items:center;gap:10px;margin:0 0 22px 2px}
  .logo{width:30px;height:30px;border-radius:8px;background:var(--brand);
    display:grid;place-items:center;color:#fff;font-family:"Fraunces",serif;
    font-weight:600;font-size:18px;box-shadow:0 4px 12px -4px rgba(227,91,38,.6)}
  .brandname{font-weight:600;letter-spacing:.02em;font-size:15px}
  .brandname span{color:var(--brand)}
  .panel{background:var(--panel);border:1px solid var(--line);border-radius:20px;
    padding:40px 40px 30px;box-shadow:var(--shadow)}
  .eyebrow{font-size:12px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;
    color:var(--brand);margin:0 0 12px}
  h1{font-family:"Fraunces",serif;font-weight:600;font-size:32px;line-height:1.15;
    margin:0 0 10px;letter-spacing:-.01em}
  .lede{color:var(--muted);font-size:15.5px;margin:0 0 26px;max-width:44ch}
  .drop{position:relative;border:1.5px dashed #d8cfbf;border-radius:14px;
    padding:38px 24px;text-align:center;cursor:pointer;transition:border-color .18s,background .18s;
    background:linear-gradient(180deg,#fcfaf6,#faf7f1)}
  .drop:hover{border-color:var(--brand)}
  .drop.drag{border-color:var(--brand);background:#fdf1ea}
  .drop.ready{border-style:solid;border-color:var(--ok);background:#f1f8f3}
  .drop input{position:absolute;inset:0;opacity:0;cursor:pointer}
  .ic{width:42px;height:42px;margin:0 auto 12px;color:var(--brand);display:block}
  .drop.ready .ic{color:var(--ok)}
  .dtitle{font-weight:600;font-size:15.5px}
  .dhint{color:var(--muted);font-size:13px;margin-top:5px}
  .fname{margin-top:12px;font-weight:600;color:var(--ok);font-size:14px;word-break:break-all}
  .swap{margin-top:6px;font-size:12.5px;color:var(--muted);text-decoration:underline}
  .btn{margin-top:20px;width:100%;padding:15px;border:0;border-radius:12px;
    font-family:"Inter";font-size:15.5px;font-weight:600;cursor:pointer;
    background:var(--brand);color:#fff;transition:background .18s,transform .05s;
    box-shadow:0 10px 22px -12px rgba(227,91,38,.85)}
  .btn:hover:not(:disabled){background:var(--brand-deep)}
  .btn:active:not(:disabled){transform:translateY(1px)}
  .btn:disabled{background:#e4ddd0;color:#a79f90;cursor:not-allowed;box-shadow:none}
  .prog{margin-top:20px;display:none}
  .prog.show{display:block}
  .ptrack{height:9px;border-radius:99px;background:#ece6da;overflow:hidden}
  .pfill{height:100%;width:0%;border-radius:99px;
    background:linear-gradient(90deg,var(--brand),#f2954f);transition:width .3s ease}
  .pmsg{margin-top:9px;font-size:13px;color:var(--muted);display:flex;justify-content:space-between}
  .done{margin-top:20px;display:none;align-items:center;gap:10px;
    background:#f1f8f3;border:1px solid #cfe7d8;border-radius:12px;padding:14px 16px}
  .done.show{display:flex}
  .done .tick{width:22px;height:22px;color:var(--ok);flex:none}
  .done .txt{font-size:14px}
  .done .txt b{display:block;color:var(--ink)}
  .done .txt span{color:var(--muted)}
  .err{margin-top:18px;display:none;background:#fdecea;border:1px solid #f3c4bd;
    color:#a3271f;border-radius:10px;padding:12px 14px;font-size:14px}
  .err.show{display:block}
  .foot{margin:18px 4px 0;color:var(--muted);font-size:12.5px;text-align:center}
  .legend{margin-top:22px;padding-top:18px;border-top:1px solid var(--line);
    display:flex;gap:18px;flex-wrap:wrap;font-size:12.5px;color:var(--muted)}
  .legend .dot{display:inline-block;width:11px;height:11px;border-radius:3px;
    background:#fbe6c9;border:1px solid #edc98f;margin-right:6px;vertical-align:-1px}
  @media (max-width:520px){ .panel{padding:30px 22px 24px} h1{font-size:26px} }
  @media (prefers-reduced-motion:reduce){ .pfill{transition:none} }
</style>
</head>
<body>
  <div class="wrap">
    <div class="brandrow">
      <div class="logo">D</div>
      <div class="brandname">Decor<span>pot</span></div>
    </div>

    <div class="panel">
      <p class="eyebrow">Initiation BOM</p>
      <h1>Turn a quotation into a ready BOM.</h1>
      <p class="lede">Upload the ERP quotation export. Get back the import-ready
        Initiation BOM — Lists, Price and BOM sheets, rooms in quotation order.</p>

      <label class="drop" id="drop">
        <svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
        </svg>
        <div class="dtitle">Choose CSV file or drag it here</div>
        <div class="dhint">ERP quotation export · .csv</div>
        <div class="fname" id="fname"></div>
        <div class="swap" id="swap" style="display:none">choose a different file</div>
        <input type="file" id="csv" accept=".csv">
      </label>

      <button class="btn" id="btn" disabled>Generate Initiation BOM</button>

      <div class="prog" id="prog">
        <div class="ptrack"><div class="pfill" id="pfill"></div></div>
        <div class="pmsg"><span id="pmsg">Starting…</span><span id="ppct">0%</span></div>
      </div>

      <div class="done" id="done">
        <svg class="tick" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>
        <div class="txt"><b>Initiation BOM ready</b><span id="donesub">Your download has started.</span></div>
      </div>

      <div class="err" id="err"></div>

      <div class="legend">
        <div><span class="dot"></span>Highlighted rows need a quick check against the working drawing (e.g. False Ceiling area, Wallpaper code).</div>
      </div>
    </div>

    <p class="foot">Master data built in · one CSV at a time · nothing is stored</p>
  </div>

<script>
  const inp=document.getElementById('csv'), drop=document.getElementById('drop'),
        fname=document.getElementById('fname'), swap=document.getElementById('swap'),
        btn=document.getElementById('btn'), prog=document.getElementById('prog'),
        pfill=document.getElementById('pfill'), pmsg=document.getElementById('pmsg'),
        ppct=document.getElementById('ppct'), done=document.getElementById('done'),
        donesub=document.getElementById('donesub'), err=document.getElementById('err');

  let file=null;
  function setFile(f){ file=f; if(!f) return;
    fname.textContent=f.name; swap.style.display='block';
    drop.classList.add('ready'); btn.disabled=false;
    err.classList.remove('show'); done.classList.remove('show');
  }
  inp.addEventListener('change',e=>{ if(inp.files.length) setFile(inp.files[0]); });
  ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag');}));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag');}));
  drop.addEventListener('drop',e=>{ if(e.dataTransfer.files.length){ inp.files=e.dataTransfer.files; setFile(e.dataTransfer.files[0]); }});

  const steps=[[12,'Reading quotation…'],[34,'Matching line items…'],
               [58,'Applying rules & merging…'],[80,'Building Lists · Price · BOM…'],[93,'Finalising workbook…']];

  btn.addEventListener('click',()=>{
    if(!file) return;
    btn.disabled=true; done.classList.remove('show'); err.classList.remove('show');
    prog.classList.add('show'); pfill.style.width='0%'; ppct.textContent='0%'; pmsg.textContent='Starting…';

    let i=0;
    const tick=setInterval(()=>{ if(i<steps.length){ pfill.style.width=steps[i][0]+'%';
      ppct.textContent=steps[i][0]+'%'; pmsg.textContent=steps[i][1]; i++; } },550);

    const fd=new FormData(); fd.append('csv',file);
    fetch('/generate',{method:'POST',body:fd})
      .then(r=>{ if(!r.ok) return r.json().then(j=>{throw new Error(j.error||'Generation failed')});
        return r.blob().then(b=>({b,r})); })
      .then(({b,r})=>{ clearInterval(tick);
        pfill.style.width='100%'; ppct.textContent='100%'; pmsg.textContent='Done';
        const cd=r.headers.get('Content-Disposition')||'';
        const m=cd.match(/filename="?([^"]+)"?/); const name=m?m[1]:'Initiation_BOM.xlsx';
        const url=URL.createObjectURL(b); const a=document.createElement('a');
        a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
        donesub.textContent=name+' downloaded.';
        setTimeout(()=>{ prog.classList.remove('show'); done.classList.add('show');
          btn.disabled=false; btn.textContent='Generate again'; },500);
      })
      .catch(e=>{ clearInterval(tick); prog.classList.remove('show');
        err.textContent=e.message+' — please check the CSV and try again.'; err.classList.add('show');
        btn.disabled=false;
      });
  });
</script>
</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(PAGE)

@app.route("/generate", methods=["POST"])
def generate():
    f = request.files.get("csv")
    if not f or f.filename == "":
        return jsonify({"error": "Please choose a CSV file"}), 400
    if not f.filename.lower().endswith(".csv"):
        return jsonify({"error": "That isn't a .csv file"}), 400

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input.csv")
        out_path = os.path.join(tmp, "Initiation_BOM.xlsx")
        f.save(in_path)
        try:
            cfg = ib.load_config(MASTER)
            results, unmatched = ib.process_csv(in_path, cfg)
            ib.write_output(results, unmatched, out_path, template=TEMPLATE)
            with open(out_path, "rb") as fh:
                data = fh.read()
        except Exception as e:
            return jsonify({"error": f"Could not process this file: {e}"}), 500

    base = os.path.splitext(os.path.basename(f.filename))[0]
    return send_file(io.BytesIO(data),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=f"Initiation_BOM_{base}.xlsx")

@app.route("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
