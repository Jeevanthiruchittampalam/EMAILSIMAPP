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
DISPLAY_LIMIT = 500  # rows to render in the table for performance (raised a bit for UX)

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
    * { margin:0; padding:0; box-sizing:border-box; }
    :root{
      --brand:#667eea;
      --brand2:#764ba2;
      --ok:#4caf50;
      --ink:#333;
      --muted:#666;
      --bg:#f9faff;
      --chip:#e8f5e8;
      --chiptext:#1b5e20;
    }
    body { font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 100%); min-height:100vh; padding:20px; color:var(--ink); }
    .container { max-width:1400px; margin:0 auto; }
    .header { text-align:center; color:#fff; margin-bottom:30px; }
    .header h1 { font-size:2.5em; font-weight:300; margin-bottom:10px; }
    .main-panel { background:rgba(255,255,255,.96); border-radius:20px; box-shadow:0 20px 40px rgba(0,0,0,.1); overflow:hidden; backdrop-filter:blur(10px); }
    .upload-section { padding:40px; border-bottom:1px solid #e0e0e0; }
    .upload-zone { border:3px dashed var(--brand); border-radius:15px; padding:40px; text-align:center; background:#f8f9ff; transition:.3s; cursor:pointer; }
    .upload-zone:hover { background:#f0f2ff; border-color:#5a6fd8; transform:translateY(-2px); }
    .upload-zone.dragover { background:#e8ebff; border-color:#4c5dd6; }
    .upload-icon { font-size:3em; color:var(--brand); margin-bottom:15px; }
    .config-section { padding:24px 40px; background:var(--bg); }
    .keyword-input { display:flex; gap:10px; margin-bottom:16px; }
    .keyword-input input { flex:1; padding:12px 15px; border:2px solid #e0e0e0; border-radius:8px; font-size:14px; }
    .keyword-input button { padding:12px 20px; background:var(--brand); color:#fff; border:none; border-radius:8px; cursor:pointer; font-weight:600; }
    .keyword-tags { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; min-height:30px; }
    .keyword-tag { background:#e8ebff; color:#4c5dd6; padding:6px 12px; border-radius:20px; font-size:13px; display:flex; align-items:center; gap:6px; }
    .keyword-tag .remove { cursor:pointer; font-weight:bold; color:#ff6b6b; }
    .process-btn { width:100%; padding:15px; background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 100%); color:#fff; border:none; border-radius:10px; font-size:16px; font-weight:600; cursor:pointer; transition:transform .3s; }
    .process-btn:hover { transform:translateY(-2px); }
    .process-btn:disabled { opacity:.6; cursor:not-allowed; transform:none; }
    .results-section { padding:20px 30px 30px; display:none; }
    .results-header { display:flex; gap:12px; align-items:center; flex-wrap:wrap; margin-bottom:14px; }
    .results-stats { background:#e8f5e8; color:var(--chiptext); padding:10px 14px; border-radius:10px; border-left:4px solid var(--ok); font-weight:600; }
    .toolbar { margin-left:auto; display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
    .toolbar input[type="search"]{ padding:10px 12px; border:1px solid #ddd; border-radius:8px; min-width:240px; }
    .toolbar select, .toolbar button { padding:10px 12px; border-radius:8px; border:1px solid #ddd; background:#fff; cursor:pointer; }
    .toolbar .download { background:var(--ok); color:#fff; border:none; }
    .toolbar .csv { background:#009688; color:#fff; border:none; }
    .results-table { background:#fff; border-radius:10px; overflow:auto; box-shadow:0 4px 15px rgba(0,0,0,.08); max-height:65vh; }
    table { width:100%; border-collapse:separate; border-spacing:0; min-width:1400px; }
    th, td { padding:10px 12px; text-align:left; border-bottom:1px solid #eee; max-width:520px; word-wrap:break-word; white-space:pre-wrap; vertical-align:top; background:#fff; }
    thead th { background:var(--brand); color:#fff; position:sticky; top:0; z-index:5; user-select:none; cursor:pointer; }
    thead th.sortable:hover { filter:brightness(0.95); }
    thead th .sort-ind { font-size:12px; opacity:0.9; margin-left:6px; }
    tr:nth-child(even) td { background:#fafbff; }
    tr:hover td { background:#f5f7ff; }
    .chip { background:var(--chip); color:var(--chiptext); padding:2px 6px; border-radius:10px; font-size:11px; }
    .loading { display:none; text-align:center; padding:20px; color:var(--brand); }
    .spinner { border:3px solid #f3f3f3; border-top:3px solid var(--brand); border-radius:50%; width:30px; height:30px; animation:spin 1s linear infinite; margin:0 auto 15px; }
    @keyframes spin { 0%{transform:rotate(0)} 100%{transform:rotate(360deg)} }
    .section-title { font-size:1.2em; font-weight:700; color:#333; margin-bottom:10px; }
    .default-keywords { background:#fff3cd; padding:12px 14px; border-radius:8px; margin-bottom:12px; border-left:4px solid #ffc107; }
    .default-keywords h4 { color:#856404; margin-bottom:6px; font-size:14px; }
    .default-keywords ul { list-style:none; display:flex; flex-wrap:wrap; gap:8px; }
    .default-keywords li { background:#fff; padding:4px 8px; border-radius:15px; font-size:12px; color:#856404; }
    .alert { padding:12px; margin:0 40px 20px; border-radius:8px; display:none; }
    .alert.success { background:#d4edda; color:#155724; border:1px solid #c3e6cb; }
    .alert.error { background:#f8d7da; color:#721c24; border:1px solid #f5c6cb; }
    mark { padding:0 2px; border-radius:3px; }

    /* Sticky first two columns for readability (subject + body/first col) */
    td.sticky-1, th.sticky-1 { position:sticky; left:0; z-index:4; }
    td.sticky-2, th.sticky-2 { position:sticky; left:280px; z-index:4; }
    th.sticky-1, th.sticky-2 { background:#5568d5; }
    td.sticky-1, td.sticky-2 { background:#fff; }
    /* Adjust the default width for first two columns to keep them readable */
    .col-0 { width:280px; max-width:280px; }
    .col-1 { width:560px; max-width:560px; }
    /* Modal for full cell content */
    .modal-backdrop { position:fixed; inset:0; background:rgba(0,0,0,.35); display:none; align-items:center; justify-content:center; padding:24px; }
    .modal { width:min(1000px, 90vw); max-height:80vh; overflow:auto; background:#fff; border-radius:12px; padding:18px; box-shadow:0 20px 50px rgba(0,0,0,.25); }
    .modal h3 { margin-bottom:10px; }
    .modal pre { white-space:pre-wrap; word-wrap:break-word; font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; background:#f8f9fb; border:1px solid #eee; padding:12px; border-radius:8px; }
    .modal .close { float:right; border:none; background:#eee; border-radius:8px; padding:6px 10px; cursor:pointer; }
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
          <p>Supports .xlsx and .xls files (up to {{ max_mb }} MB)</p>
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
          <input type="text" id="keywordInput" placeholder="Add additional keywords... (press Enter)" />
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

          <div class="toolbar">
            <input type="search" id="globalSearch" placeholder="Search results…" />
            <label>
              Rows / page
              <select id="pageSize">
                <option value="25">25</option>
                <option value="50" selected>50</option>
                <option value="100">100</option>
                <option value="200">200</option>
              </select>
            </label>
            <button class="download" onclick="downloadResults()">📥 XLSX</button>
            <button class="csv" onclick="downloadCsv()">⬇️ CSV</button>
          </div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; margin:8px 0 12px;">
          <div id="pagination" style="font-size:14px; color:var(--muted);"></div>
          <div id="columnToggles" style="display:flex; gap:8px; flex-wrap:wrap;"></div>
        </div>

        <div class="results-table" id="resultsTable"></div>
      </div>
    </div>
  </div>

  <!-- Modal -->
  <div class="modal-backdrop" id="modalBackdrop" onclick="hideModal(event)">
    <div class="modal" onclick="event.stopPropagation()">
      <button class="close" onclick="hideModal()">Close</button>
      <h3 id="modalTitle"></h3>
      <pre id="modalBody"></pre>
    </div>
  </div>

  <script>
    // ------- State -------
    let additionalKeywords = [];
    let currentFileName = '';
    let resultsData = null;

    // Client-side table state
    let headers = [];
    let rawRows = [];        // Full rows from server (truncated to DISPLAY_LIMIT on server)
    let filteredRows = [];   // After global search
    let sortState = { index: null, dir: 1 }; // 1 asc, -1 desc
    let currentPage = 1;
    let pageSize = 50;
    let hiddenCols = new Set(); // column indexes hidden by toggles

    // ------- Upload handling -------
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

    function showAlert(message,type){
      const a=document.getElementById('alertBox');
      a.className=`alert ${type}`; a.textContent=message; a.style.display='block';
      setTimeout(()=>{a.style.display='none';},5000);
    }

    // ------- Keyword management -------
    function addKeyword(){ const input=document.getElementById('keywordInput'); const keyword=input.value.trim().toLowerCase(); if(keyword && !additionalKeywords.includes(keyword)){ additionalKeywords.push(keyword); updateKeywordTags(); input.value=''; } }
    function removeKeyword(k){ additionalKeywords=additionalKeywords.filter(x=>x!==k); updateKeywordTags(); }
    function updateKeywordTags(){ const c=document.getElementById('keywordTags'); c.innerHTML=additionalKeywords.map(k=>`<div class="keyword-tag">${escapeHtml(k)}<span class="remove" onclick="removeKeyword('${k.replace(/'/g, "\\'")}')">&times;</span></div>`).join(''); }
    document.getElementById('keywordInput').addEventListener('keypress', e=>{ if(e.key==='Enter'){ addKeyword(); }});

    // ------- Processing -------
    function processFile(){
      if(!currentFileName){ showAlert('Please upload a file first','error'); return; }
      document.getElementById('loading').style.display='block';
      document.getElementById('resultsSection').style.display='none';
      fetch('/process', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ filename: currentFileName, additional_keywords: additionalKeywords })
      })
      .then(r=>r.json())
      .then(data=>{
        document.getElementById('loading').style.display='none';
        if(data.success){
          resultsData = data;
          headers = data.headers;
          rawRows = data.results;
          filteredRows = rawRows.slice();
          sortState = { index:null, dir:1 };
          currentPage = 1;
          pageSize = parseInt(document.getElementById('pageSize').value,10);
          initColumnToggles();
          render();
          showAlert(`Processing complete: ${data.matching_count} emails found`, 'success');
        } else {
          showAlert('Processing failed: ' + data.error, 'error');
        }
      })
      .catch(err=>{
        document.getElementById('loading').style.display='none';
        showAlert('Processing failed: '+err.message,'error');
      });
    }

    // ------- Rendering / Table utilities -------
    function escapeHtml(s){
      return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    function getVisibleHeaders(){
      return headers.map((h, idx) => ({h, idx})).filter(obj => !hiddenCols.has(obj.idx));
    }

    function render(){
      const resultsSection=document.getElementById('resultsSection');
      const resultsStats=document.getElementById('resultsStats');
      const resultsTable=document.getElementById('resultsTable');

      resultsStats.innerHTML = `<strong>📊 Results Found:</strong> ${resultsData.matching_count} matching emails out of ${resultsData.total_count} total — showing up to {{ display_limit }} rows.`;

      if(!filteredRows.length){
        resultsTable.innerHTML='<div style="padding:40px; text-align:center; color:#666;">No matching emails found.</div>';
        resultsSection.style.display='block';
        document.getElementById('pagination').textContent = '';
        return;
      }

      // Pagination slice
      const total = filteredRows.length;
      const pages = Math.max(1, Math.ceil(total / pageSize));
      if(currentPage > pages) currentPage = pages;
      const start = (currentPage-1)*pageSize;
      const end = Math.min(total, start+pageSize);
      const slice = filteredRows.slice(start, end);

      // Build header
      const visible = getVisibleHeaders();
      const ths = visible.map((obj, i) => {
        const colIdx = obj.idx;
        const label = escapeHtml(obj.h);
        const sticky = colIdx===0 ? 'sticky-1 col-0' : (colIdx===1 ? 'sticky-2 col-1' : '');
        const isSorted = sortState.index === colIdx;
        const ind = isSorted ? (sortState.dir===1 ? '▲' : '▼') : '';
        return `<th class="sortable ${sticky}" data-col="${colIdx}">${label}<span class="sort-ind">${ind}</span></th>`;
      }).join('') + `<th>Match Reason</th>`;

      // Build body
      const tds = slice.map(row => {
        const cells = visible.map((obj) => {
          const colIdx = obj.idx;
          const raw = String(row[headers[colIdx]] || '');
          const isBody = /body|message|content/i.test(headers[colIdx]);
          const title = escapeHtml(raw);
          const display = raw.length > (isBody ? 800 : 160) ? escapeHtml(raw.substring(0, isBody?800:160) + '…') : escapeHtml(raw);
          const sticky = colIdx===0 ? 'sticky-1' : (colIdx===1 ? 'sticky-2' : '');
          // open modal on click for full view
          return `<td class="${sticky}" title="${title}" onclick="openModal('${escapeHtml(headers[colIdx])}', \`${title}\`)">${display}</td>`;
        }).join('');
        return `<tr>${cells}<td><span class="chip">${escapeHtml(row._match_reason || '')}</span></td></tr>`;
      }).join('');

      resultsTable.innerHTML = `
        <table>
          <thead><tr>${ths}</tr></thead>
          <tbody>${tds}</tbody>
        </table>
      `;

      // Bind sort handlers
      resultsTable.querySelectorAll('thead th.sortable').forEach(th=>{
        th.onclick = ()=>{
          const colIdx = parseInt(th.getAttribute('data-col'),10);
          if(sortState.index === colIdx){ sortState.dir *= -1; }
          else { sortState.index = colIdx; sortState.dir = 1; }
          stableSort(colIdx, sortState.dir);
          currentPage = 1;
          render();
        };
      });

      // Pagination UI
      document.getElementById('pagination').innerHTML = `
        <span>Showing <strong>${start+1}-${end}</strong> of <strong>${total}</strong></span>
        &nbsp;|&nbsp;
        <button onclick="firstPage()" ${currentPage===1?'disabled':''}>⏮</button>
        <button onclick="prevPage()" ${currentPage===1?'disabled':''}>◀</button>
        <span>Page ${currentPage} / ${pages}</span>
        <button onclick="nextPage()" ${currentPage===pages?'disabled':''}>▶</button>
        <button onclick="lastPage()" ${currentPage===pages?'disabled':''}>⏭</button>
      `;

      resultsSection.style.display='block';
    }

    function stableSort(colIdx, dir){
      const key = headers[colIdx];
      // decorate-sort-undecorate for stability
      const decorated = filteredRows.map((row, i)=>({i, row, v: String(row[key] || '').toLowerCase()}));
      decorated.sort((a,b)=>{
        if(a.v < b.v) return -1*dir;
        if(a.v > b.v) return 1*dir;
        return a.i - b.i; // stable
      });
      filteredRows = decorated.map(d=>d.row);
    }

    // ------- Search & Pagination -------
    document.getElementById('globalSearch').addEventListener('input', function(){
      const q = this.value.trim().toLowerCase();
      if(!q){
        filteredRows = rawRows.slice();
      } else {
        filteredRows = rawRows.filter(row=>{
          return headers.some(h => String(row[h]||'').toLowerCase().includes(q));
        });
      }
      currentPage = 1;
      render();
    });

    document.getElementById('pageSize').addEventListener('change', function(){
      pageSize = parseInt(this.value,10);
      currentPage = 1;
      render();
    });

    function firstPage(){ currentPage = 1; render(); }
    function prevPage(){ if(currentPage>1){ currentPage--; render(); } }
    function nextPage(){ const pages = Math.ceil(filteredRows.length / pageSize)||1; if(currentPage<pages){ currentPage++; render(); } }
    function lastPage(){ currentPage = Math.ceil(filteredRows.length / pageSize)||1; render(); }

    // ------- Column toggles -------
    function initColumnToggles(){
      const holder = document.getElementById('columnToggles');
      holder.innerHTML = '<span style="color:#555; font-weight:600;">Columns:</span>';
      headers.forEach((h, idx)=>{
        const id = 'col_'+idx;
        const checked = !hiddenCols.has(idx) ? 'checked' : '';
        const disabled = (idx===0 || idx===1) ? 'disabled' : ''; // keep first two visible (sticky)
        holder.innerHTML += `
          <label style="display:flex; align-items:center; gap:6px; font-size:13px;">
            <input type="checkbox" id="${id}" ${checked} ${disabled} onchange="toggleCol(${idx}, this.checked)" />
            ${escapeHtml(h)}
          </label>`;
      });
    }
    function toggleCol(idx, show){
      if(!show) hiddenCols.add(idx);
      else hiddenCols.delete(idx);
      render();
    }

    // ------- Modal -------
    function openModal(title, body){
      document.getElementById('modalTitle').textContent = title;
      document.getElementById('modalBody').textContent = body;
      document.getElementById('modalBackdrop').style.display = 'flex';
    }
    function hideModal(ev){ document.getElementById('modalBackdrop').style.display='none'; }

    // ------- Downloads -------
    function downloadResults(){
      if(!resultsData || resultsData.matching_count===0){ showAlert('No results to download','error'); return; }
      window.location.href = '/download';
    }
    function downloadCsv(){
      if(!resultsData || resultsData.matching_count===0){ showAlert('No results to download','error'); return; }
      window.location.href = '/download_csv';
    }
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
    return render_template_string(
        HTML_TEMPLATE,
        default_keywords=KEYWORDS,
        important_senders=IMPORTANT_SENDERS,
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
    """Return XLSX in-memory (works well on Windows/macOS/Render)."""
    try:
        if "filtered_data" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})

        rows = []
        for row in processed_data["filtered_data"]:
            rows.append({k: v for k, v in row.items() if k != "_match_reason"})
        df = pd.DataFrame(rows)

        # Build Excel in memory to avoid filesystem issues
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="results")
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"EMAILSIM_output_{timestamp}.xlsx"
        return send_file(output, as_attachment=True, download_name=out_name)

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/download_csv")
def download_csv():
    """Optional CSV download."""
    try:
        if "filtered_data" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})

        rows = []
        for row in processed_data["filtered_data"]:
            rows.append({k: v for k, v in row.items() if k != "_match_reason"})
        df = pd.DataFrame(rows)

        from io import StringIO, BytesIO
        csv_buf = StringIO()
        df.to_csv(csv_buf, index=False)
        byte_buf = BytesIO(csv_buf.getvalue().encode("utf-8"))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"EMAILSIM_output_{timestamp}.csv"
        return send_file(byte_buf, as_attachment=True, download_name=out_name, mimetype="text/csv")
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

# expose the Flask app for gunicorn
application = app
