#!/usr/bin/env python3
"""
Decorpot - Initiation BOM Generator (web app)
Upload an ERP quotation CSV, get the Initiation BOM Excel back.
Master data is built in; team only uploads the CSV.
"""
import os
import io
import tempfile
from flask import Flask, request, send_file, render_template_string, flash, redirect

import initiation_bom as ib

app = Flask(__name__)
app.secret_key = "decorpot-initiation-bom"

BASE = os.path.dirname(os.path.abspath(__file__))
MASTER = os.path.join(BASE, "master_data.xlsx")
TEMPLATE = os.path.join(BASE, "bom_template.xlsx")

PAGE = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Initiation BOM Generator · Decorpot</title>
<style>
  :root{
    --ink:#1c1c1a; --muted:#6b6a64; --line:#e2e0d8;
    --paper:#faf9f6; --card:#ffffff; --brand:#F26A30; --brand-ink:#303030;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
    background:var(--paper);color:var(--ink);line-height:1.6;
    min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:16px;
    max-width:560px;width:100%;padding:40px;box-shadow:0 1px 3px rgba(0,0,0,.04)}
  .badge{display:inline-block;font-size:12px;font-weight:600;letter-spacing:.08em;
    text-transform:uppercase;color:var(--brand);margin-bottom:8px}
  h1{font-size:26px;margin:0 0 6px;color:var(--brand-ink)}
  p.sub{color:var(--muted);margin:0 0 28px;font-size:15px}
  .drop{border:2px dashed var(--line);border-radius:12px;padding:36px 20px;text-align:center;
    transition:.2s;cursor:pointer;background:var(--paper)}
  .drop:hover{border-color:var(--brand)}
  .drop.has{border-color:var(--brand);background:#fff6f1}
  .drop input{display:none}
  .drop .icon{font-size:32px}
  .drop .hint{color:var(--muted);font-size:14px;margin-top:6px}
  .fname{font-weight:600;color:var(--brand-ink);margin-top:8px;word-break:break-all}
  button{margin-top:20px;width:100%;padding:14px;border:0;border-radius:10px;
    background:var(--brand-ink);color:#fff;font-size:15px;font-weight:600;cursor:pointer;transition:.2s}
  button:hover{background:#000}
  button:disabled{opacity:.4;cursor:not-allowed}
  .flash{background:#fdecea;color:#a3271f;border:1px solid #f5c1bb;border-radius:8px;
    padding:12px 14px;margin-bottom:18px;font-size:14px}
  .foot{color:var(--muted);font-size:12.5px;margin-top:22px;text-align:center}
</style>
</head>
<body>
  <div class="card">
    <span class="badge">Decorpot</span>
    <h1>Initiation BOM Generator</h1>
    <p class="sub">Upload the ERP quotation CSV. Download the ready-to-import Initiation BOM.</p>
    {% with msgs = get_flashed_messages() %}
      {% if msgs %}{% for m in msgs %}<div class="flash">{{ m }}</div>{% endfor %}{% endif %}
    {% endwith %}
    <form method="post" action="/generate" enctype="multipart/form-data" id="f">
      <label class="drop" id="drop">
        <div class="icon">📄</div>
        <div><strong>Choose CSV file</strong> or drag it here</div>
        <div class="hint">ERP quotation export (.csv)</div>
        <div class="fname" id="fname"></div>
        <input type="file" name="csv" id="csv" accept=".csv" required>
      </label>
      <button type="submit" id="btn" disabled>Generate Initiation BOM</button>
    </form>
    <div class="foot">Master data is built in · rooms grouped in quotation order</div>
  </div>
<script>
  const inp=document.getElementById('csv'),drop=document.getElementById('drop'),
        fname=document.getElementById('fname'),btn=document.getElementById('btn');
  function show(){if(inp.files.length){fname.textContent=inp.files[0].name;
    drop.classList.add('has');btn.disabled=false}}
  inp.addEventListener('change',show);
  ['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('has')}));
  ['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();}));
  drop.addEventListener('drop',ev=>{ev.preventDefault();if(ev.dataTransfer.files.length){inp.files=ev.dataTransfer.files;show()}});
  document.getElementById('f').addEventListener('submit',()=>{btn.textContent='Generating…';btn.disabled=true});
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
        flash("Please choose a CSV file.")
        return redirect("/")
    if not f.filename.lower().endswith(".csv"):
        flash("That doesn't look like a .csv file. Please upload the ERP quotation CSV.")
        return redirect("/")

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
            flash(f"Could not process this file: {e}")
            return redirect("/")

    base = os.path.splitext(os.path.basename(f.filename))[0]
    return send_file(io.BytesIO(data),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=f"Initiation_BOM_{base}.xlsx")

@app.route("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
