import os
import re
import json
import tempfile
import threading
import webbrowser
from datetime import datetime

import pandas as pd
from flask import Flask, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename

# ------------------------------
# Configuration
# ------------------------------
KEYWORDS = [
    "investment analysis",
    "exit plan",
    "financing thind",  # kept as-is in case it's intentional
    "takeout",
    "sale of highline",
    "tower c",
    "seville will be paid back",
]

IMPORTANT_SENDERS = ["kingset", "abacus north"]

MAX_FILE_MB = 16
DISPLAY_LIMIT = 200  # rows to render in the table for performance

# ------------------------------
# Flask setup
# ------------------------------
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = tempfile.gettempdir()
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_MB * 1024 * 1024

processed_data: dict[str, object] = {}

# ------------------------------
# HTML (kept close to your original, with a few small tweaks)
# ------------------------------
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Email Filter & Analysis Tool</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); min-height:100vh; padding:20px; }
    .container { max-width:1200px; margin:0 auto; }
    .header { text-align:center; color:#fff; margin-bottom:30px; }
    .header h1 { font-size:2.5em; font-weight:300; margin-bottom:10px; }
    .main-panel { background:rgba(255,255,255,.95); border-radius:20px; box-shadow:0 20px 40px rgba(0,0,0,.1); overflow:hidden; backdrop-filter:blur(10px); }
    .upload-section { padding:40px; border-bottom:1px solid #e0e0e0; }
    .upload-zone { border:3px dashed #667eea; border-radius:15px; padding:40px; text-align:center; background:#f8f9ff; transition:.3s; cursor:pointer; }
    .upload-zone:hover { background:#f0f2ff; border-color:#5a6fd8; transform:translateY(-2px); }
    .upload-zone.dragover { background:#e8ebff; border-color:#4c5dd6; }
    .upload-icon { font-size:3em; color:#667eea; margin-bottom:15px; }
    .config-section { padding:30px 40px; background:#f9faff; }
    .keyword-input { display:flex; gap:10px; margin-bottom:20px; }
    .keyword-input input { flex:1; padding:12px 15px; border:2px solid #e0e0e0; border-radius:8px; font-size:14px; }
    .keyword-input button { padding:12px 20px; background:#667eea; color:#fff; border:none; border-radius:8px; cursor:pointer; font-weight:600; }
    .keyword-tags { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px; min-height:30px; }
    .keyword-tag { background:#e8ebff; color:#4c5dd6; padding:6px 12px; border-radius:20px; font-size:13px; display:flex; align-items:center; gap:6px; }
    .keyword-tag .remove { cursor:pointer; font-weight:bold; color:#ff6b6b; }
    .process-btn { width:100%; padding:15px; background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; border:none; border-radius:10px; font-size:16px; font-weight:600; cursor:pointer; transition:transform .3s; }
    .process-btn:hover { transform:translateY(-2px); }
    .process-btn:disabled { opacity:.6; cursor:not-allowed; transform:none; }
    .results-section { padding:30px 40px; display:none; }
    .results-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; flex-wrap:wrap; gap:15px; }
    .results-stats { background:#e8f5e8; padding:15px 20px; border-radius:10px; border-left:4px solid #4caf50; }
    .download-btn { padding:10px 20px; background:#4caf50; color:#fff; border:none; border-radius:8px; cursor:pointer; font-weight:600; }
    .results-table { background:#fff; border-radius:10px; overflow:hidden; box-shadow:0 4px 15px rgba(0,0,0,.1); max-height:500px; overflow-y:auto; }
    table { width:100%; border-collapse:collapse; }
    th, td { padding:12px 15px; text-align:left; border-bottom:1px solid #e0e0e0; max-width:300px; word-wrap:break-word; white-space:pre-wrap; vertical-align:top; }
    th { background:#667eea; color:#fff; font-weight:600; position:sticky; top:0; z-index:10; }
    tr:hover { background:#f8f9ff; }
    .loading { display:none; text-align:center; padding:20px; color:#667eea; }
    .spinner { border:3px solid #f3f3f3; border-top:3px solid #667eea; border-radius:50%; width:30px; height:30px; animation:spin 1s linear infinite; margin:0 auto 15px; }
    @keyframes spin { 0%{transform:rotate(0)} 100%{transform:rotate(360deg)} }
    .section-title { font-size:1.3em; font-weight:600; color:#333; margin-bottom:15px; }
    .default-keywords { background:#fff3cd; padding:15px; border-radius:8px; margin-bottom:20px; border-left:4px solid #ffc107; }
    .default-keywords h4 { color:#856404; margin-bottom:8px; }
    .default-keywords ul { list-style:none; display:flex; flex-wrap:wrap; gap:8px; }
    .default-keywords li { background:#fff; padding:4px 8px; border-radius:15px; font-size:12px; color:#856404; }
    .alert { padding:15px; margin-bottom:20px; border-radius:8px; display:none; }
    .alert.success { background:#d4edda; color:#155724; border:1px solid #c3e6cb; }
    .alert.error { background:#f8d7da; color:#721c24; border:1px solid #f5c6cb; }
    mark { padding:0 2px; border-radius:3px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📧 Email Filter & Analysis Tool</h1>
      <p>Upload Excel files and filter emails by keywords and important senders</p>
    </div>

    <div class="main-panel">
      <div class="alert" id="alertBox"></div>

      <!-- Upload Section -->
      <div class="upload-section">
        <div class="upload-zone" id="uploadZone">
          <div class="upload-icon">📁</div>
          <h3>Drop your Excel file here or click to browse</h3>
          <p>Supports .xlsx and .xls files</p>
          <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" />
        </div>
      </div>

      <!-- Configuration Section -->
      <div class="config-section">
        <div class="section-title">🔧 Search Configuration</div>
        <div class="default-keywords">
          <h4>Default Keywords & Senders:</h4>
          <ul>
            {% for keyword in default_keywords %}
            <li>{{ keyword }}</li>
            {% endfor %}
            {% for sender in important_senders %}
            <li>{{ sender }} (sender)</li>
            {% endfor %}
          </ul>
        </div>

        <div class="keyword-input">
          <input type="text" id="keywordInput" placeholder="Add additional keywords..." />
          <button onclick="addKeyword()">Add Keyword</button>
        </div>
        <div class="keyword-tags" id="keywordTags"></div>
        <button class="process-btn" id="processBtn" onclick="processFile()" disabled>🚀 Process File</button>
      </div>

      <!-- Loading Section -->
      <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>Processing your file...</p>
      </div>

      <!-- Results Section -->
      <div class="results-section" id="resultsSection">
        <div class="results-header">
          <div class="results-stats" id="resultsStats"></div>
          <button class="download-btn" onclick="downloadResults()">📥 Download Results</button>
        </div>
        <div class="results-table" id="resultsTable"></div>
      </div>
    </div>
  </div>

  <script>
    let additionalKeywords = [];
    let currentFileName = '';
    let resultsData = null;

    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');

    uploadZone.addEventListener('click', () => fileInput.click());
    uploadZone.addEventListener('dragover', handleDragOver);
    uploadZone.addEventListener('dragleave', handleDragLeave);
    uploadZone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);

    function handleDragOver(e){ e.preventDefault(); uploadZone.classList.add('dragover'); }
    function handleDragLeave(e){ e.preventDefault(); uploadZone.classList.remove('dragover'); }
    function handleDrop(e){ e.preventDefault(); uploadZone.classList.remove('dragover'); const files=e.dataTransfer.files; if(files.length>0){ uploadFile(files[0]); } }
    function handleFileSelect(e){ if(e.target.files.length>0){ uploadFile(e.target.files[0]); } }

    function uploadFile(file){
      if(!file.name.match(/\.(xlsx|xls)$/i)){ showAlert('Please select an Excel file (.xlsx or .xls)','error'); return; }
      const formData=new FormData(); formData.append('file', file);
      fetch('/upload', { method:'POST', body:formData })
        .then(r=>r.json())
        .then(data=>{
          if(data.success){
            currentFileName=data.filename;
            uploadZone.innerHTML=`<div class="upload-icon">✅</div><h3>File loaded successfully!</h3><p>${file.name} (${data.rows} rows)</p>`;
            document.getElementById('processBtn').disabled=false;
            showAlert(`File uploaded successfully: ${data.rows} rows loaded`, 'success');
          } else { showAlert('Error uploading file: '+data.error,'error'); }
        })
        .catch(err=> showAlert('Upload failed: '+ err.message,'error'));
    }

    function showAlert(message,type){ const a=document.getElementById('alertBox'); a.className=`alert ${type}`; a.textContent=message; a.style.display='block'; setTimeout(()=>{a.style.display='none';},5000); }

    function addKeyword(){ const input=document.getElementById('keywordInput'); const keyword=input.value.trim().toLowerCase(); if(keyword && !additionalKeywords.includes(keyword)){ additionalKeywords.push(keyword); updateKeywordTags(); input.value=''; } }
    function removeKeyword(k){ additionalKeywords=additionalKeywords.filter(x=>x!==k); updateKeywordTags(); }
    function updateKeywordTags(){ const c=document.getElementById('keywordTags'); c.innerHTML=additionalKeywords.map(k=>`<div class="keyword-tag">${k}<span class="remove" onclick="removeKeyword('${k}')">&times;</span></div>`).join(''); }

    document.getElementById('keywordInput').addEventListener('keypress', e=>{ if(e.key==='Enter'){ addKeyword(); }});

    function processFile(){
      if(!currentFileName){ showAlert('Please upload a file first','error'); return; }
      document.getElementById('loading').style.display='block';
      document.getElementById('resultsSection').style.display='none';
      fetch('/process', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ filename: currentFileName, additional_keywords: additionalKeywords }) })
        .then(r=>r.json())
        .then(data=>{
          document.getElementById('loading').style.display='none';
          if(data.success){ resultsData=data; displayResults(data); showAlert(`Processing complete: ${data.matching_count} emails found`,'success'); }
          else { showAlert('Processing failed: '+data.error,'error'); }
        })
        .catch(err=>{ document.getElementById('loading').style.display='none'; showAlert('Processing failed: '+err.message,'error'); });
    }

    function displayResults(data){
      const resultsSection=document.getElementById('resultsSection');
      const resultsStats=document.getElementById('resultsStats');
      const resultsTable=document.getElementById('resultsTable');
      resultsStats.innerHTML = `<strong>📊 Results Found:</strong> ${data.matching_count} matching emails out of ${data.total_count} total`;

      if(data.matching_count===0){ resultsTable.innerHTML='<div style="padding:40px; text-align:center; color:#666;">No matching emails found.</div>'; }
      else {
        const headers = data.headers;
        const rows = data.results.slice(0, {{ display_limit }});
        const tableHTML = `
          <table>
            <thead>
              <tr>
                ${headers.map(h=>`<th>${h}</th>`).join('')}
                <th>Match Reason</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(row=>`
                <tr>
                  ${headers.map(h=>{
                    const raw = String(row[h] || '');
                    const isBody = /body|message|content/i.test(h);
                    const maxLen = isBody ? 500 : 100;
                    const display = raw.length > maxLen ? raw.substring(0, maxLen) + '…' : raw;
                    const title = raw.replace(/"/g, '&quot;');
                    return `<td title="${title}">${display}</td>`;
                  }).join('')}
                  <td><span style="background:#e8f5e8; padding:2px 6px; border-radius:10px; font-size:11px;">${row._match_reason}</span></td>
                </tr>`).join('')}
            </tbody>
          </table>`;
        resultsTable.innerHTML = tableHTML;
      }
      resultsSection.style.display='block';
    }

    function downloadResults(){ if(!resultsData || resultsData.matching_count===0){ showAlert('No results to download','error'); return; } window.location.href='/download'; }
  </script>
</body>
</html>
"""

# ------------------------------
# Helpers
# ------------------------------

def _clean_text(text: object) -> str:
    """Clean up text by removing Excel artifacts and extra whitespace."""
    if pd.isna(text) or text == "":
        return ""
    s = str(text)
    # Excel/HTML artifacts
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


def _phrase_in_text(text: str, phrases: list[str]) -> bool:
    if not text:
        return False
    lo = text.lower()
    for p in phrases:
        pat = r"\b" + re.escape(p) + r"\b"
        if re.search(pat, lo):
            return True
    return False

# ------------------------------
# Routes
# ------------------------------

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE, default_keywords=KEYWORDS, important_senders=IMPORTANT_SENDERS, display_limit=DISPLAY_LIMIT)


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

        filename = secure_filename(f.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_{filename}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        f.save(path)

        # Try reading with openpyxl first; fallback to xlrd for .xls
        try:
            if filename.lower().endswith(".xlsx"):
                df = pd.read_excel(path, engine="openpyxl")
            else:
                # For .xls you need xlrd==1.2.0
                df = pd.read_excel(path, engine="xlrd")
        except Exception:
            # Fallback to default engine resolution
            df = pd.read_excel(path)

        df = df.fillna("")
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].apply(_clean_text)

        processed_data["current_file"] = path
        processed_data["original_data"] = df

        return jsonify({"success": True, "filename": filename, "rows": int(len(df))})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/process", methods=["POST"])
def process_file():
    try:
        data = request.json or {}
        extra = data.get("additional_keywords", [])
        if "original_data" not in processed_data:
            return jsonify({"success": False, "error": "No file uploaded"})

        df: pd.DataFrame = processed_data["original_data"]  # type: ignore
        all_kw = [k.lower() for k in (KEYWORDS + list(extra)) if k]
        matches: list[dict] = []

        sender_cols = [c for c in df.columns if "from" in c.lower() or "to" in c.lower()]
        imp_senders_lower = [s.lower() for s in IMPORTANT_SENDERS]

        for _, row in df.iterrows():
            # Join all text for keyword scan
            full_text = " ".join(str(row.get(c, "")) for c in df.columns)

            if _phrase_in_text(full_text, all_kw):
                rd = {k: _clean_text(v) if pd.notna(v) else "" for k, v in row.to_dict().items()}
                rd["_match_reason"] = "Keyword Match"
                matches.append(rd)
                continue

            # Sender scan
            participants = [str(row.get(c, "")) for c in sender_cols]
            if any(any(s in p.lower() for s in imp_senders_lower) for p in participants):
                rd = {k: _clean_text(v) if pd.notna(v) else "" for k, v in row.to_dict().items()}
                rd["_match_reason"] = "Important Sender"
                matches.append(rd)

        processed_data["filtered_data"] = matches
        processed_data["headers"] = list(df.columns)

        return jsonify({
            "success": True,
            "total_count": int(len(df)),
            "matching_count": int(len(matches)),
            "results": matches[:DISPLAY_LIMIT],
            "headers": list(df.columns),
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/download")
def download_results():
    try:
        if "filtered_data" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})

        rows = []
        for row in processed_data["filtered_data"]:
            rows.append({k: v for k, v in row.items() if k != "_match_reason"})
        df = pd.DataFrame(rows)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"EMAILSIM_output_{timestamp}.xlsx"
        out_path = os.path.join(app.config["UPLOAD_FOLDER"], out_name)
        df.to_excel(out_path, index=False)

        return send_file(out_path, as_attachment=True, download_name=out_name)

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ------------------------------
# Dev server bootstrap
# ------------------------------

def _open_browser_delayed(url: str):
    import time
    time.sleep(1.5)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main():
    print("🚀 Starting Email Filter Web Application…")
    print("📂 Upload folder:", app.config["UPLOAD_FOLDER"])  # noqa: T201

    # Pick an open port (try 5000, then 5001)
    import socket
    port = 5000
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", port)) == 0:
            print(f"⚠️  Port {port} in use. Trying 5001…")  # noqa: T201
            port = 5001

    url = f"http://127.0.0.1:{port}/"
    print(f"🌐 Starting web server on: {url}")  # noqa: T201

    t = threading.Thread(target=_open_browser_delayed, args=(url,), daemon=True)
    t.start()

    try:
        app.run(debug=True, host="127.0.0.1", port=port, use_reloader=False)
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")  # noqa: T201


if __name__ == "__main__":
    main()
