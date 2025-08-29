import os
import re
import tempfile
import threading
import traceback
import webbrowser
from datetime import datetime

import pandas as pd
from flask import Flask, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename

# ------------------------------
# Configuration
# ------------------------------
# Important senders are folded into defaults as plain terms you can remove/re-add.
DEFAULT_KEYWORDS = [
    "investment analysis",
    "exit plan",
    "financing thind",
    "takeout",
    "sale of highline",
    "tower c",
    "seville will be paid back",
    # (former IMPORTANT_SENDERS)
    "kingset",
    "abacus north",
]

MAX_FILE_MB = 16
DISPLAY_LIMIT = 500  # rows to render in the table

# ------------------------------
# Flask setup
# ------------------------------
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = tempfile.gettempdir()
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_MB * 1024 * 1024

processed_data: dict[str, object] = {}

# ------------------------------
# HTML + CSS + JS
# ------------------------------
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Email Filter & Analysis Tool</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  :root{
    --brand:#667eea;--brand2:#764ba2;--accent:#00bcd4;
    --ink:#222;--muted:#666;--panel:#fff;--divider:#e6e8f0;--grid:#e9ecf5;
    --shadow:0 0 0 1px rgba(0,0,0,.03),0 6px 18px rgba(0,0,0,.06);
    --ok:#4caf50;--warn:#ff6b6b;--chipbg:#eef1ff;--chipbd:#e0e5ff;
  }
  body{font-family:Segoe UI,Tahoma,Geneva,Verdana,sans-serif;background:linear-gradient(135deg,var(--brand),var(--brand2));min-height:100vh;color:var(--ink);padding:20px}
  .container{max-width:1440px;margin:0 auto}
  .header{color:#fff;text-align:center;margin-bottom:22px}
  .header h1{font-weight:300;letter-spacing:.2px}

  .main{background:#fff;border-radius:18px;box-shadow:0 20px 40px rgba(0,0,0,.1);overflow:hidden}

  .section{padding:24px 36px;background:var(--panel)}
  .section+.section{border-top:1px solid var(--divider)}
  .sub{font-weight:800;font-size:13px;color:#444;letter-spacing:.12em;text-transform:uppercase;margin-bottom:10px}

  /* Upload */
  .upload{border:3px dashed var(--brand);border-radius:14px;padding:36px;text-align:center;background:#f8f9ff;cursor:pointer;transition:.25s}
  .upload:hover{background:#f0f2ff}
  .upload .icon{font-size:2.2rem;color:var(--brand);margin-bottom:8px}

  /* Config */
  .grid{display:grid;grid-template-columns:1.25fr .75fr;gap:20px}
  .card{border:1px solid var(--divider);border-radius:12px;padding:14px;box-shadow:var(--shadow)}
  .card h4{font-size:14px;text-transform:uppercase;letter-spacing:.08em;color:#444;margin-bottom:8px}
  .desc{font-size:13px;color:#555}
  .kw-input{display:flex;gap:8px;margin-top:10px}
  .kw-input input{flex:1;padding:10px 12px;border:2px solid #e0e0e0;border-radius:8px}
  .btn{padding:10px 14px;border-radius:8px;border:1px solid #d7d9e0;background:#fff;cursor:pointer}
  .btn.primary{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;border:none;font-weight:700}
  .btn.green{background:var(--ok);color:#fff;border:none}
  .chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
  .chip{display:inline-flex;gap:6px;align-items:center;background:var(--chipbg);border:1px solid var(--chipbd);border-radius:18px;padding:6px 10px;font-size:13px}
  .chip .x{color:var(--warn);font-weight:700;cursor:pointer}
  .chip.add{background:#fff;border-style:dashed;cursor:pointer}

  .toggle{display:flex;align-items:center;gap:10px;margin-top:8px}
  .switch{position:relative;width:56px;height:30px}
  .switch input{opacity:0;width:0;height:0}
  .slider{position:absolute;inset:0;background:#dfe3f7;border:1px solid var(--grid);border-radius:30px;cursor:pointer;transition:.2s}
  .slider:before{content:"";position:absolute;width:24px;height:24px;left:3px;top:2px;background:#fff;border-radius:50%;box-shadow:0 2px 6px rgba(0,0,0,.12);transition:.25s}
  .switch input:checked + .slider{background:#c6f7d2;border-color:#b6eabf}
  .switch input:checked + .slider:before{transform:translateX(26px)}

  .alert{display:none;margin:0 36px 16px;border-radius:8px;padding:10px}
  .alert.success{background:#d4edda;color:#155724;border:1px solid #c3e6cb}
  .alert.error{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb}

  /* Results toolbar */
  .row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
  .row.push{justify-content:space-between;margin:10px 0}
  .stats{background:#eef8ff;border-left:4px solid var(--accent);border-radius:8px;padding:8px 12px;font-weight:700;color:#0a4a6b}
  .toolbar input[type=search]{padding:10px 12px;border:1px solid #d7d9e0;border-radius:8px;min-width:220px}

  /* Table: header sticky only; body columns scroll like Excel */
  .tablewrap{border:1px solid var(--divider);border-radius:12px;box-shadow:var(--shadow);overflow:auto;max-height:66vh;background:#fff}
  table{border-collapse:separate;border-spacing:0;min-width:1400px;width:100%}
  thead th{position:sticky;top:0;background:var(--brand);color:#fff;z-index:3;padding:10px 12px}
  th,td{border-bottom:1px solid var(--grid);border-right:1px solid var(--grid);vertical-align:top}
  th:last-child,td:last-child{border-right:none}
  tbody td{padding:10px 12px}
  tbody tr:nth-child(even) td{background:#fbfbff}
  .cell{white-space:pre-wrap;word-wrap:break-word;max-height:240px;overflow:auto}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>📧 Email Filter & Analysis Tool</h1>
    <p>Upload and filter by keywords (ANY or ALL). Remove or add terms freely.</p>
  </div>

  <div class="main">

    <div id="alert" class="alert"></div>

    <!-- Upload -->
    <div class="section">
      <div class="sub">Upload</div>
      <div id="uploadZone" class="upload">
        <div class="icon">📁</div>
        <div>Drop your Excel here or click to browse</div>
        <small>Supports .xlsx and .xls (up to {{ max_mb }} MB)</small>
        <input id="fileInput" type="file" accept=".xlsx,.xls" style="display:none" />
      </div>
    </div>

    <!-- Config -->
    <div class="section">
      <div class="sub">Configuration</div>
      <div class="grid">
        <div class="card">
          <h4>Keywords</h4>
          <div class="desc">Active terms (click × to remove). Add new below. If no terms remain, no rows will match.</div>
          <div id="activeChips" class="chips"></div>

          <div class="kw-input">
            <input id="kwInput" placeholder="Add keyword… (Enter)" />
            <button class="btn" onclick="addKeyword()">Add</button>
          </div>

          <div class="desc" style="margin-top:10px">Available defaults (click to add back):</div>
          <div id="defaultChips" class="chips"></div>
        </div>

        <div class="card">
          <h4>Match Mode</h4>
          <div class="toggle">
            <label class="switch">
              <input id="allToggle" type="checkbox" />
              <span class="slider"></span>
            </label>
            <div><strong id="modeLabel">ANY terms (default)</strong><br/><small class="muted">ON = ALL terms required</small></div>
          </div>
          <div style="margin-top:12px">
            <button id="processBtn" class="btn primary" onclick="processFile()" disabled>🚀 Process</button>
          </div>
        </div>
      </div>
    </div>

    <!-- Results -->
    <div id="resultsSec" class="section" style="display:none">
      <div class="sub">Results</div>

      <div class="row">
        <div id="stats" class="stats"></div>
        <div class="row toolbar">
          <input id="searchBox" type="search" placeholder="Search results…" />
          <label>Rows<select id="pageSize" class="btn"><option>25</option><option selected>50</option><option>100</option><option>200</option></select></label>
          <button class="btn" onclick="clearSearch()">Clear</button>
          <button class="btn green" onclick="downloadXlsx()">📥 XLSX</button>
          <button class="btn" onclick="downloadCsv()">⬇️ CSV</button>
        </div>
      </div>

      <div class="row push">
        <div id="pager" class="muted"></div>
        <div id="colToggles" class="row"></div>
      </div>

      <div id="tableWrap" class="tablewrap"></div>
    </div>
  </div>
</div>

<script>
  // ---------- state ----------
  const defaults = {{ default_keywords|tojson }};
  let active = [...defaults];            // terms to send to server
  let currentFile = '';
  let allHeaders = [];
  let rawRows = [];
  let filtRows = [];
  let page = 1;
  let per = 50;
  let sort = { col: null, dir: 1 };
  let hidden = new Set();

  const alertBox = document.getElementById('alert');

  // ---------- helpers ----------
  function showAlert(msg, type='success'){
    alertBox.className = 'alert ' + type;
    alertBox.textContent = msg;
    alertBox.style.display = 'block';
    setTimeout(()=>alertBox.style.display='none', 6000);
  }
  function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function toJsonSafe(resp){
    return resp.text().then(t=>{
      const ct = resp.headers.get('content-type')||'';
      if(ct.includes('application/json')){ return JSON.parse(t); }
      throw new Error('Server returned non-JSON:\n' + t.slice(0,1500));
    });
  }

  // ---------- chips ----------
  function drawChips(){
    const activeEl = document.getElementById('activeChips');
    activeEl.innerHTML = active.map(k=>(
      `<span class="chip">${esc(k)} <span class="x" onclick="removeTerm('${k.replace(/'/g,"\\'")}')">×</span></span>`
    )).join('') || '<div class="desc" style="margin-top:6px">No active terms.</div>';

    const defaultEl = document.getElementById('defaultChips');
    const pool = defaults.filter(k=>!active.includes(k));
    defaultEl.innerHTML = pool.map(k=>(
      `<span class="chip add" onclick="addDefault('${k.replace(/'/g,"\\'")}')">+ ${esc(k)}</span>`
    )).join('') || '<div class="desc" style="margin-top:6px">All defaults active.</div>';
  }
  function removeTerm(k){ active = active.filter(x=>x!==k); drawChips(); }
  function addDefault(k){ if(!active.includes(k)) active.push(k); drawChips(); }
  function addKeyword(){
    const inp = document.getElementById('kwInput'); const v = (inp.value||'').trim().toLowerCase();
    if(v && !active.includes(v)){ active.push(v); }
    inp.value=''; drawChips();
  }
  document.getElementById('kwInput').addEventListener('keydown', e=>{ if(e.key==='Enter') addKeyword(); });
  drawChips();

  // ---------- upload ----------
  const zone = document.getElementById('uploadZone');
  const fileInput = document.getElementById('fileInput');
  zone.addEventListener('click', ()=>fileInput.click());
  zone.addEventListener('dragover', e=>{ e.preventDefault(); zone.classList.add('hover'); });
  zone.addEventListener('dragleave', e=>{ e.preventDefault(); zone.classList.remove('hover'); });
  zone.addEventListener('drop', e=>{ e.preventDefault(); zone.classList.remove('hover'); if(e.dataTransfer.files.length){ doUpload(e.dataTransfer.files[0]); }});
  fileInput.addEventListener('change', e=>{ if(e.target.files.length){ doUpload(e.target.files[0]); }});

  function doUpload(file){
    if(!/\.(xlsx|xls)$/i.test(file.name)){ showAlert('Please choose .xlsx or .xls','error'); return; }
    const fd = new FormData(); fd.append('file', file);
    fetch('/upload',{method:'POST',body:fd})
      .then(toJsonSafe)
      .then(j=>{
        if(!j.success) throw new Error(j.error||'Upload failed');
        currentFile = j.filename;
        document.getElementById('processBtn').disabled = false;
        zone.innerHTML = `<div class="icon">✅</div><div>Loaded ${esc(file.name)} (${j.rows} rows)</div>`;
        showAlert('File uploaded');
      })
      .catch(err=>showAlert(err.message,'error'));
  }

  // ---------- mode label ----------
  const allToggle = document.getElementById('allToggle');
  const modeLabel = document.getElementById('modeLabel');
  allToggle.addEventListener('change', ()=>{ modeLabel.textContent = allToggle.checked ? 'ALL terms required' : 'ANY terms (default)'; });

  // ---------- process ----------
  function processFile(){
    if(!currentFile){ showAlert('Upload a file first.','error'); return; }
    const payload = {
      filename: currentFile,
      additional_keywords: active.slice(),  // exact active terms
      require_all: !!allToggle.checked
    };
    fetch('/process', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)})
      .then(toJsonSafe)
      .then(j=>{
        if(!j.success) throw new Error(j.error||'Process failed');
        allHeaders = j.headers;
        rawRows = j.results;
        filtRows = rawRows.slice();
        page = 1;
        per = parseInt(document.getElementById('pageSize').value,10);
        hidden = new Set();
        sort = {col:null, dir:1};
        drawToggles();
        drawTable();
        document.getElementById('resultsSec').style.display='block';
        document.getElementById('stats').innerHTML = `Matches: <strong>${j.matching_count}</strong> / ${j.total_count} (showing up to {{ display_limit }})`;
        showAlert('Processing complete.');
      })
      .catch(err=>showAlert(err.message,'error'));
  }

  // ---------- table render ----------
  function visibleHeaders(){ return allHeaders.map((h,i)=>({h,i})).filter(o=>!hidden.has(o.i)); }
  function drawToggles(){
    const wrap = document.getElementById('colToggles');
    wrap.innerHTML = '<strong style="margin-right:6px">Columns:</strong>';
    allHeaders.forEach((h,i)=>{
      const id='c_'+i;
      wrap.insertAdjacentHTML('beforeend',
        `<label class="row" style="gap:6px"><input id="${id}" type="checkbox" ${hidden.has(i)?'':'checked'} onchange="toggleCol(${i}, this.checked)" /> ${esc(h)}</label>`
      );
    });
  }
  function toggleCol(i, show){ show ? hidden.delete(i) : hidden.add(i); drawTable(); }

  function drawTable(){
    const wrap = document.getElementById('tableWrap');
    if(!filtRows.length){ wrap.innerHTML = '<div style="padding:24px;color:#666;text-align:center">No matching rows.</div>'; drawPager(0,0,0); return; }

    const total=filtRows.length, pages=Math.max(1,Math.ceil(total/per));
    if(page>pages) page=pages;
    const s=(page-1)*per, e=Math.min(total,s+per);
    const vis = visibleHeaders();

    const thead = `<thead><tr>${vis.map(v=>`<th>${esc(v.h)}</th>`).join('')}<th>Match Reason</th></tr></thead>`;
    const tbody = `<tbody>` + filtRows.slice(s,e).map(row=>{
      const tds = vis.map(v=>{
        const val = String(row[allHeaders[v.i]] ?? '');
        return `<td><div class="cell">${esc(val)}</div></td>`;
      }).join('');
      return `<tr>${tds}<td><div class="cell">${esc(row._match_reason||'')}</div></td></tr>`;
    }).join('') + `</tbody>`;

    wrap.innerHTML = `<table>${thead}${tbody}</table>`;
    drawPager(s+1,e,total);

    // sort handlers
    wrap.querySelectorAll('thead th').forEach((th,idx)=>{
      th.onclick = ()=>{
        const colIdx = idx; // includes only visible cols; map back:
        let realIdx = (colIdx===vis.length) ? null : vis[colIdx]?.i; // last header is "Match Reason"
        if(realIdx==null) return; // don't sort by match reason
        if(sort.col===realIdx){ sort.dir*=-1; } else { sort.col=realIdx; sort.dir=1; }
        const key = allHeaders[realIdx];
        filtRows = filtRows
          .map((r,i)=>({r,i,v:String(r[key]||'').toLowerCase()}))
          .sort((a,b)=> a.v<b.v ? -1*sort.dir : a.v>b.v ? 1*sort.dir : a.i-b.i)
          .map(o=>o.r);
        page=1;
        drawTable();
      };
    });
  }

  function drawPager(start,end,total){
    document.getElementById('pager').innerHTML =
      total ? `Showing <strong>${start}-${end}</strong> of <strong>${total}</strong>` : '';
  }

  // search + page size
  document.getElementById('searchBox').addEventListener('input', function(){
    const q = this.value.trim().toLowerCase();
    filtRows = q ? rawRows.filter(r=>allHeaders.some(h=>String(r[h]||'').toLowerCase().includes(q))) : rawRows.slice();
    page=1; drawTable();
  });
  document.getElementById('pageSize').addEventListener('change', function(){ per=parseInt(this.value,10); page=1; drawTable(); });
  function clearSearch(){ const b=document.getElementById('searchBox'); if(!b.value) return; b.value=''; filtRows=rawRows.slice(); page=1; drawTable(); }

  // downloads
  function downloadXlsx(){ if(!rawRows.length){ showAlert('Nothing to download','error'); return; } window.location='/download'; }
  function downloadCsv(){ if(!rawRows.length){ showAlert('Nothing to download','error'); return; } window.location='/download_csv'; }
</script>
</body>
</html>
"""

# ------------------------------
# Helpers
# ------------------------------
def _clean_text(x) -> str:
    if pd.isna(x) or x == "":
        return ""
    s = str(x)
    s = (
        s.replace("_x000D_", " ")
        .replace("*x000D*", " ")
        .replace("_x000A_", " ")
        .replace("*x000A*", " ")
        .replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&#39;", "'")
    )
    s = re.sub(r"_\s+_", " ", s)
    s = re.sub(r"\s+_\s+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _phrase_in_text(text: str, phrase: str) -> bool:
    if not text:
        return False
    lo = text.lower()
    pat = r"\b" + re.escape(phrase) + r"\b"
    return re.search(pat, lo) is not None

def _matches(full_text: str, phrases: list[str], require_all: bool) -> bool:
    phrases = [p for p in phrases if isinstance(p, str) and p.strip()]
    if not phrases:
        return False
    return all(_phrase_in_text(full_text, p) for p in phrases) if require_all \
           else any(_phrase_in_text(full_text, p) for p in phrases)

def _json_error(message: str, code: int = 500):
    tb = traceback.format_exc(limit=3)
    resp = jsonify({"success": False, "error": message, "trace": tb})
    resp.status_code = code
    return resp

# ------------------------------
# Routes
# ------------------------------
@app.route("/")
def index():
    return render_template_string(
        HTML_TEMPLATE,
        default_keywords=DEFAULT_KEYWORDS,
        display_limit=DISPLAY_LIMIT,
        max_mb=MAX_FILE_MB,
    )

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        if "file" not in request.files:
            return jsonify({"success": False, "error": "No file selected"})
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"success": False, "error": "No file selected"})
        if not f.filename.lower().endswith((".xlsx", ".xls")):
            return jsonify({"success": False, "error": "Invalid file type"})

        name = secure_filename(f.filename)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"{stamp}_{name}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], name)
        f.save(path)

        # Load
        try:
            if name.lower().endswith(".xlsx"):
                df = pd.read_excel(path, engine="openpyxl")
            else:
                df = pd.read_excel(path, engine="xlrd")
        except Exception:
            df = pd.read_excel(path)

        df = df.fillna("")
        for c in df.columns:
            if df[c].dtype == "object":
                df[c] = df[c].apply(_clean_text)

        processed_data["path"] = path
        processed_data["df"] = df

        return jsonify({"success": True, "filename": name, "rows": int(len(df))})
    except Exception as e:
        return _json_error(f"Upload failed: {e}")

@app.route("/process", methods=["POST"])
def process_file():
    try:
        data = request.get_json(silent=True) or {}
        active_terms = [str(p).lower().strip() for p in data.get("additional_keywords", []) if str(p).strip()]
        require_all = bool(data.get("require_all", False))

        if "df" not in processed_data:
            return jsonify({"success": False, "error": "No file uploaded"})

        df: pd.DataFrame = processed_data["df"]  # type: ignore

        matches = []
        # Build a full lowercased row text for scanning (fast enough for XLS volumes)
        for _, row in df.iterrows():
            full_text = " ".join(str(row.get(c, "")) for c in df.columns).lower()
            if _matches(full_text, active_terms, require_all):
                d = {k: _clean_text(v) if pd.notna(v) else "" for k, v in row.to_dict().items()}
                d["_match_reason"] = "ALL terms" if require_all else "ANY term"
                matches.append(d)

        processed_data["filtered"] = matches
        processed_data["headers"] = list(df.columns)

        return jsonify({
            "success": True,
            "total_count": int(len(df)),
            "matching_count": int(len(matches)),
            "results": matches[:DISPLAY_LIMIT],
            "headers": list(df.columns),
        })
    except Exception as e:
        return _json_error(f"Processing failed: {e}")

@app.route("/download")
def download_xlsx():
    try:
        if "filtered" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})
        rows = [{k:v for k,v in r.items() if k!="_match_reason"} for r in processed_data["filtered"]]
        df = pd.DataFrame(rows)
        from io import BytesIO
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="results")
        buf.seek(0)
        name = f"EMAILSIM_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(buf, as_attachment=True, download_name=name)
    except Exception as e:
        return _json_error(f"Download failed: {e}")

@app.route("/download_csv")
def download_csv():
    try:
        if "filtered" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})
        rows = [{k:v for k,v in r.items() if k!="_match_reason"} for r in processed_data["filtered"]]
        df = pd.DataFrame(rows)
        from io import StringIO, BytesIO
        s = StringIO(); df.to_csv(s, index=False)
        b = BytesIO(s.getvalue().encode("utf-8"))
        name = f"EMAILSIM_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        return send_file(b, as_attachment=True, download_name=name, mimetype="text/csv")
    except Exception as e:
        return _json_error(f"CSV export failed: {e}")

# ------------------------------
# Dev server bootstrap
# ------------------------------
def _open_browser(url:str):
    import time; time.sleep(1.2)
    try: webbrowser.open(url)
    except Exception: pass

def main():
    print("🚀 Starting Email Filter Web Application…")
    print("📂 Upload folder:", app.config["UPLOAD_FOLDER"])
    import socket
    port=5000
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", port)) == 0: port=5001
    url=f"http://127.0.0.1:{port}/"
    threading.Thread(target=_open_browser, args=(url,), daemon=True).start()
    app.run(debug=True, host="127.0.0.1", port=port, use_reloader=False)

if __name__ == "__main__":
    main()

# expose for gunicorn
application = app
