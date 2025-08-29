import os
import re
import json
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
DEFAULT_KEYWORDS = [
    "investment analysis",
    "exit plan",
    "financing thind",  # kept as-is in case it's intentional
    "takeout",
    "sale of highline",
    "tower c",
    "seville will be paid back",
]

MAX_FILE_MB = 16
DISPLAY_LIMIT = 500  # rows to render in the table for performance

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
      --brand:#667eea; --brand2:#764ba2; --ok:#4caf50; --ink:#222; --muted:#666;
      --panel:#ffffff; --divider:#e6e8f0; --tableGrid:#e9ecf5; --accent:#00bcd4;
      --warn:#ff6b6b; --highlight:#eef1ff; --shadow:0 0 0 1px rgba(0,0,0,.03), 0 6px 18px rgba(0,0,0,.06);
    }
    body { font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 100%); min-height:100vh; padding:20px; color:var(--ink); }
    .container { max-width:1440px; margin:0 auto; }
    .header { text-align:center; color:#fff; margin-bottom:28px; }
    .header h1 { font-size:2.6em; font-weight:300; margin-bottom:8px; letter-spacing:.2px; }
    .main-panel { background:rgba(255,255,255,.96); border-radius:20px; box-shadow:0 20px 40px rgba(0,0,0,.1); overflow:hidden; backdrop-filter:blur(10px); }
    .section { padding:26px 40px; background:var(--panel); }
    .section + .section { border-top:1px solid var(--divider); }
    .subhead { font-weight:800; font-size:14px; color:#4a4a4a; letter-spacing:.12em; text-transform:uppercase; margin-bottom:12px; }

    /* Upload */
    .upload-zone { border:3px dashed var(--brand); border-radius:15px; padding:40px; text-align:center; background:#f8f9ff; transition:.3s; cursor:pointer; }
    .upload-zone:hover { background:#f0f2ff; border-color:#5a6fd8; transform:translateY(-2px); }
    .upload-zone.dragover { background:#e8ebff; border-color:#4c5dd6; }
    .upload-icon { font-size:3em; color:var(--brand); margin-bottom:15px; }

    /* Config */
    .config-wrap { display:grid; grid-template-columns: 1.3fr .7fr; gap:22px; }
    .card { background:#fff; border:1px solid var(--divider); border-radius:12px; padding:16px; box-shadow:var(--shadow); }
    .card h4 { margin-bottom:10px; font-size:15px; letter-spacing:.08em; text-transform:uppercase; color:#444; }
    .desc { color:#555; font-size:13px; margin-top:4px; }
    .keyword-input { display:flex; gap:10px; margin-top:8px; }
    .keyword-input input { flex:1; padding:12px 15px; border:2px solid #e0e0e0; border-radius:8px; font-size:14px; }
    .keyword-input button { padding:12px 20px; background:var(--brand); color:#fff; border:none; border-radius:8px; cursor:pointer; font-weight:600; }
    .chiplist { display:flex; flex-wrap:wrap; gap:8px; margin-top:12px; }
    .chip { display:inline-flex; align-items:center; gap:8px; background:var(--highlight); color:#3b4cca; padding:6px 12px; border-radius:20px; font-size:13px; border:1px solid #e0e5ff; }
    .chip .remove { cursor:pointer; color:var(--warn); font-weight:700; }
    .chip.add-back { background:#fff; color:#3b4cca; border-style:dashed; cursor:pointer; }

    /* Mode / actions */
    .mode-toggle { display:flex; align-items:center; gap:10px; margin-top:8px; }
    .mode-toggle .switch { position:relative; display:inline-block; width:56px; height:30px; }
    .mode-toggle .switch input{ opacity:0; width:0; height:0; }
    .mode-toggle .slider{ position:absolute; cursor:pointer; inset:0; background:#dfe3f7; transition:.2s; border-radius:30px; border:1px solid var(--tableGrid); }
    .mode-toggle .slider:before{ position:absolute; content:""; height:24px; width:24px; left:3px; top:2px; background:white; transition:.25s; border-radius:50%; box-shadow:0 2px 6px rgba(0,0,0,.12); }
    .mode-toggle input:checked + .slider{ background:#c6f7d2; border-color:#b6eabf;}
    .mode-toggle input:checked + .slider:before{ transform:translateX(26px);}
    .mode-label{ font-size:13px; color:#333; }
    .actions { display:flex; align-items:center; gap:12px; margin-top:14px; }
    .process-btn { padding:13px 18px; background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 100%); color:#fff; border:none; border-radius:10px; font-size:15px; font-weight:700; cursor:pointer; transition:transform .2s; }
    .process-btn:hover { transform:translateY(-1px); }
    .process-btn:disabled { opacity:.6; cursor:not-allowed; transform:none; }

    /* Alerts */
    .alert { padding:12px; margin:0 40px 20px; border-radius:8px; display:none; }
    .alert.success { background:#d4edda; color:#155724; border:1px solid #c3e6cb; }
    .alert.error { background:#f8d7da; color:#721c24; border:1px solid #f5c6cb; }

    /* Results header & toolbar */
    .results-header { display:flex; gap:12px; align-items:center; flex-wrap:wrap; margin-bottom:10px; }
    .results-stats { background:#eef8ff; color:#0a4a6b; padding:10px 14px; border-radius:10px; border-left:4px solid var(--accent); font-weight:700; }
    .toolbar { margin-left:auto; display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
    .toolbar input[type="search"]{ padding:10px 12px; border:1px solid #d7d9e0; border-radius:8px; min-width:240px; }
    .toolbar select, .toolbar button { padding:10px 12px; border-radius:8px; border:1px solid #d7d9e0; background:#fff; cursor:pointer; }
    .toolbar .download { background:var(--ok); color:#fff; border:none; }
    .toolbar .csv { background:#009688; color:#fff; border:none; }
    .toolbar .clear { background:#fff3f3; color:#c62828; border-color:#ffebee; }

    /* Results table (no sticky columns) */
    .results-table { background:#fff; border-radius:12px; overflow:auto; box-shadow:var(--shadow); max-height:66vh; border:1px solid var(--divider); }
    table { width:100%; border-collapse:separate; border-spacing:0; min-width:1280px; }
    th, td { padding:0; border-bottom:1px solid var(--tableGrid); border-right:1px solid var(--tableGrid); vertical-align:top; background:#fff; }
    th:last-child, td:last-child { border-right:none; }
    thead th { background:var(--brand); color:#fff; position:sticky; top:0; z-index:2; user-select:none; cursor:pointer; letter-spacing:.02em; padding:10px 12px; }
    thead th.sortable:hover { filter:brightness(0.95); }
    thead th .sort-ind { font-size:12px; opacity:0.9; margin-left:6px; }
    tbody tr:nth-child(even) td { background:#fbfbff; }
    .cell { padding:10px 12px; max-height:220px; overflow:auto; white-space:pre-wrap; word-wrap:break-word; }
    .cell.small { max-height:none; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📧 Seville Email Analysis Tool 📊</h1>
      <p>Upload Excel files and filter emails by keywords</p>
    </div>

    <div class="main-panel">
      <div class="alert" id="alertBox"></div>

      <!-- Upload -->
      <div class="section">
        <div class="subhead">Upload</div>
        <div class="upload-zone" id="uploadZone">
          <div class="upload-icon">📁</div>
          <h3>Drop your Excel file here or click to browse</h3>
          <p>Supports .xlsx and .xls files (up to {{ max_mb }} MB)</p>
          <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" />
        </div>
      </div>

      <!-- Configuration -->
      <div class="section">
        <div class="subhead">Configuration</div>
        <div class="config-wrap">
          <div class="card">
            <h4>Keywords</h4>
            <div class="desc">Remove defaults with × or add your own. Only the <em>active</em> chips are searched.</div>
            <div class="keyword-input">
              <input type="text" id="keywordInput" placeholder="Add a keyword… (press Enter)" />
              <button onclick="addKeyword()">Add</button>
            </div>
            <div class="desc" style="margin-top:10px;"><strong>Active terms:</strong></div>
            <div class="chiplist" id="activeChips"></div>
            <div class="desc" style="margin-top:12px;"><strong>Available defaults:</strong> (click to add back)</div>
            <div class="chiplist" id="availableDefaults"></div>
          </div>

          <div class="card">
            <h4>Match Mode</h4>
            <div class="mode-toggle">
              <label class="switch">
                <input type="checkbox" id="requireAllToggle" />
                <span class="slider"></span>
              </label>
              <div class="mode-label"><strong id="modeLabel">ANY terms (default)</strong></div>
            </div>
            <div class="actions">
              <button class="process-btn" id="processBtn" onclick="processFile()" disabled>🚀 Process File</button>
            </div>
          </div>
        </div>
      </div>

      <!-- Loading -->
      <div class="section" id="loading" style="display:none;">
        <div class="subhead">Processing</div>
        <div class="loading">
          <div class="spinner"></div>
          <p>Crunching your spreadsheet…</p>
        </div>
      </div>

      <!-- Results -->
      <div class="section" id="resultsSection" style="display:none;">
        <div class="subhead">Results</div>
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
            <button class="clear" onclick="clearSearch()">Clear search</button>
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
    let currentFileName = '';
    const defaultKeywords = {{ default_keywords|tojson }};
    let activeKeywords = [...defaultKeywords];     // the ONLY list the backend will use

    // Client-side table state
    let headers = [];
    let rawRows = [];
    let filteredRows = [];
    let sortState = { index: null, dir: 1 };
    let currentPage = 1;
    let pageSize = 50;
    let hiddenCols = new Set();

    // Mode
    const requireAllToggle = document.getElementById('requireAllToggle');
    const modeLabel = document.getElementById('modeLabel');
    requireAllToggle.addEventListener('change', ()=>{
      modeLabel.textContent = requireAllToggle.checked ? 'ALL terms required' : 'ANY terms (default)';
    });

    // ------- Upload -------
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    uploadZone.addEventListener('click', () => fileInput.click());
    uploadZone.addEventListener('dragover', e=>{ e.preventDefault(); uploadZone.classList.add('dragover'); });
    uploadZone.addEventListener('dragleave', e=>{ e.preventDefault(); uploadZone.classList.remove('dragover'); });
    uploadZone.addEventListener('drop', e=>{ e.preventDefault(); uploadZone.classList.remove('dragover'); if(e.dataTransfer.files.length){ uploadFile(e.dataTransfer.files[0]); }});
    fileInput.addEventListener('change', e=>{ if(e.target.files.length){ uploadFile(e.target.files[0]); }});

    async function parseResponseAsJson(resp){
      const text = await resp.text();
      const ct = resp.headers.get('content-type') || '';
      if (ct.includes('application/json')) {
        try { return JSON.parse(text); }
        catch (e) { throw new Error('Bad JSON from server: ' + (e?.message || e)); }
      } else {
        throw new Error(`Non-JSON response (${resp.status}):\n` + text.slice(0, 2000));
      }
    }

    function uploadFile(file){
      if(!file.name.match(/\.(xlsx|xls)$/i)){ showAlert('Please select an Excel file (.xlsx or .xls)','error'); return; }
      const formData=new FormData(); formData.append('file', file);
      fetch('/upload', { method:'POST', body:formData })
        .then(parseResponseAsJson)
        .then(data=>{
          if(data.success){
            currentFileName=data.filename;
            uploadZone.innerHTML=`<div class="upload-icon">✅</div><h3>File loaded successfully!</h3><p>${file.name} (${data.rows} rows)</p>`;
            document.getElementById('processBtn').disabled=false;
            showAlert(`File uploaded: ${data.rows} rows`, 'success');
          } else { showAlert('Error uploading file: '+data.error,'error'); }
        })
        .catch(e=> showAlert(String(e),'error'));
    }

    function showAlert(message,type){
      const a=document.getElementById('alertBox');
      a.className=`alert ${type}`; a.textContent=message; a.style.display='block';
      setTimeout(()=>{a.style.display='none';},7000);
    }

    // ------- Keywords UI -------
    function refreshChips(){
      const act = document.getElementById('activeChips');
      act.innerHTML = activeKeywords.map(k =>
        `<span class="chip">${escapeHtml(k)} <span class="remove" title="Remove" onclick="removeActive('${k.replace(/'/g,"\\'")}')">×</span></span>`
      ).join('');

      const avail = document.getElementById('availableDefaults');
      const pool = defaultKeywords.filter(k=>!activeKeywords.includes(k));
      avail.innerHTML = pool.length
        ? pool.map(k=>`<span class="chip add-back" onclick="addBackDefault('${k.replace(/'/g,"\\'")}')">+ ${escapeHtml(k)}</span>`).join('')
        : `<span class="desc" style="padding:6px 0;">All defaults are active.</span>`;
    }
    function removeActive(k){
      const idx = activeKeywords.indexOf(k);
      if(idx>=0){ activeKeywords.splice(idx,1); refreshChips(); }
    }
    function addBackDefault(k){
      if(!activeKeywords.includes(k)){ activeKeywords.push(k); refreshChips(); }
    }
    function addKeyword(){
      const input=document.getElementById('keywordInput');
      const raw=input.value.trim();
      if(!raw) return;
      const k = raw.toLowerCase();
      if(!activeKeywords.includes(k)){ activeKeywords.push(k); }
      input.value='';
      refreshChips();
    }
    document.getElementById('keywordInput').addEventListener('keypress', e=>{ if(e.key==='Enter'){ addKeyword(); }});
    refreshChips();

    // ------- Process -------
    function processFile(){
      if(!currentFileName){ showAlert('Please upload a file first','error'); return; }
      document.getElementById('loading').style.display='block';
      document.getElementById('resultsSection').style.display='none';

      const payload = {
        filename: currentFileName,
        additional_keywords: activeKeywords,   // exact set user sees
        require_all: !!requireAllToggle.checked
      };

      fetch('/process', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) })
        .then(parseResponseAsJson)
        .then(data=>{
          document.getElementById('loading').style.display='none';
          if(!data.success){ showAlert('Processing failed: '+data.error,'error'); return; }
          headers = data.headers;
          rawRows = data.results;
          filteredRows = rawRows.slice();
          sortState = { index:null, dir:1 };
          currentPage = 1;
          pageSize = parseInt(document.getElementById('pageSize').value,10);
          initColumnToggles();
          render();
          const mode = payload.require_all ? 'ALL terms' : 'ANY terms';
          showAlert(`Done (${mode}): ${data.matching_count} matches`, 'success');
        })
        .catch(e=>{
          document.getElementById('loading').style.display='none';
          showAlert(String(e),'error');
        });
    }

    // ------- Table -------
    function escapeHtml(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
    function clearSearch(){ const box = document.getElementById('globalSearch'); if(!box.value) return; box.value=''; filteredRows=rawRows.slice(); currentPage=1; render(); }
    function getVisibleHeaders(){ return headers.map((h,i)=>({h,idx:i})).filter(o=>!hiddenCols.has(o.idx)); }

    function render(){
      const resultsSection=document.getElementById('resultsSection');
      const resultsStats=document.getElementById('resultsStats');
      const resultsTable=document.getElementById('resultsTable');

      resultsStats.innerHTML = `<strong>📊 Results:</strong> ${rawRows.length} shown (of up to {{ display_limit }}), keywords: ${activeKeywords.length ? activeKeywords.map(escapeHtml).join(', ') : '— none —'}`;

      if(!filteredRows.length){
        resultsTable.innerHTML='<div style="padding:40px; text-align:center; color:#666;">No matching rows.</div>';
        resultsSection.style.display='block';
        document.getElementById('pagination').textContent='';
        return;
      }

      const total = filteredRows.length;
      const pages = Math.max(1, Math.ceil(total / pageSize));
      if(currentPage>pages) currentPage=pages;
      const start=(currentPage-1)*pageSize, end=Math.min(total,start+pageSize);
      const slice = filteredRows.slice(start,end);

      const visible = getVisibleHeaders();
      const ths = visible.map(o=>{
        const isSorted = sortState.index===o.idx;
        const ind = isSorted ? (sortState.dir===1?'▲':'▼') : '';
        return `<th class="sortable" data-col="${o.idx}">${escapeHtml(o.h)}<span class="sort-ind">${ind}</span></th>`;
      }).join('') + `<th>Match Reason</th>`;

      const trs = slice.map(row=>{
        const cells = visible.map(o=>{
          const key=headers[o.idx]; const val=String(row[key]??'');
          const isBody = /body|message|content/i.test(key);
          return `<td title="${escapeHtml(val)}" data-col="${escapeHtml(key)}">
                    <div class="cell ${isBody?'':'small'}" ondblclick="openModalFromCell(this.parentElement)">${escapeHtml(val)}</div>
                  </td>`;
        }).join('');
        return `<tr>${cells}<td><div class="cell small"><span class="chip"> ${escapeHtml(row._match_reason||'')} </span></div></td></tr>`;
      }).join('');

      resultsTable.innerHTML = `<table><thead><tr>${ths}</tr></thead><tbody>${trs}</tbody></table>`;

      resultsTable.querySelectorAll('thead th.sortable').forEach(th=>{
        th.onclick=()=>{
          const colIdx=parseInt(th.getAttribute('data-col'),10);
          if(sortState.index===colIdx) sortState.dir*=-1; else { sortState.index=colIdx; sortState.dir=1; }
          stableSort(colIdx, sortState.dir); currentPage=1; render();
        };
      });

      document.getElementById('pagination').innerHTML =
        `<span>Showing <strong>${start+1}-${end}</strong> of <strong>${total}</strong></span>
         &nbsp;|&nbsp;
         <button onclick="firstPage()" ${currentPage===1?'disabled':''}>⏮</button>
         <button onclick="prevPage()" ${currentPage===1?'disabled':''}>◀</button>
         <span>Page ${currentPage} / ${pages}</span>
         <button onclick="nextPage()" ${currentPage===pages?'disabled':''}>▶</button>
         <button onclick="lastPage()" ${currentPage===pages?'disabled':''}>⏭</button>`;

      resultsSection.style.display='block';
    }

    function stableSort(colIdx, dir){
      const key = headers[colIdx];
      const decorated = filteredRows.map((row, i)=>({i, row, v: String(row[key]||'').toLowerCase()}));
      decorated.sort((a,b)=> (a.v<b.v?-1: a.v>b.v?1: a.i-b.i) * dir);
      filteredRows = decorated.map(d=>d.row);
    }

    document.getElementById('globalSearch').addEventListener('input', function(){
      const q=this.value.trim().toLowerCase();
      filteredRows = !q ? rawRows.slice() : rawRows.filter(r=> headers.some(h=> String(r[h]||'').toLowerCase().includes(q)));
      currentPage=1; render();
    });
    document.getElementById('pageSize').addEventListener('change', function(){ pageSize=parseInt(this.value,10); currentPage=1; render(); });
    function firstPage(){ currentPage=1; render(); }
    function prevPage(){ if(currentPage>1){ currentPage--; render(); } }
    function nextPage(){ const pages=Math.ceil(filteredRows.length/pageSize)||1; if(currentPage<pages){ currentPage++; render(); } }
    function lastPage(){ currentPage=Math.ceil(filteredRows.length/pageSize)||1; render(); }

    // Column toggles (now you can hide ANY column since nothing is sticky)
    function initColumnToggles(){
      const holder=document.getElementById('columnToggles');
      holder.innerHTML='<span style="color:#555; font-weight:700; letter-spacing:.06em; text-transform:uppercase;">Columns</span>';
      headers.forEach((h,idx)=>{
        const id='col_'+idx; const checked=!hiddenCols.has(idx)?'checked':'';
        holder.innerHTML += `
          <label style="display:flex; align-items:center; gap:6px; font-size:13px; background:#fff; padding:4px 8px; border:1px solid #e9ecf5; border-radius:8px;">
            <input type="checkbox" id="${id}" ${checked} onchange="toggleCol(${idx}, this.checked)" /> ${escapeHtml(h)}
          </label>`;
      });
    }
    function toggleCol(idx, show){ if(!show) hiddenCols.add(idx); else hiddenCols.delete(idx); render(); }

    // Modal
    function openModal(title, body){ document.getElementById('modalTitle').textContent=title; document.getElementById('modalBody').textContent=body; document.getElementById('modalBackdrop').style.display='flex'; }
    function openModalFromCell(td){ openModal(td.getAttribute('data-col')||'Details', td.getAttribute('title')||''); }
    function hideModal(){ document.getElementById('modalBackdrop').style.display='none'; }

    // Downloads
    function downloadResults(){ if(!rawRows.length){ showAlert('No results to download','error'); return; } window.location.href='/download'; }
    function downloadCsv(){ if(!rawRows.length){ showAlert('No results to download','error'); return; } window.location.href='/download_csv'; }
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
    """Return True if an exact phrase (word-boundaries) is in the text."""
    if not text:
        return False
    lo = text.lower()
    pat = r"\b" + re.escape(phrase) + r"\b"
    return re.search(pat, lo) is not None

def _matches_by_mode(full_text: str, phrases: list[str], require_all: bool) -> bool:
    """Implements ANY vs ALL mode for phrase matching."""
    phrases = [p for p in phrases if isinstance(p, str) and p.strip()]
    if not phrases:
        return False
    return all(_phrase_in_text(full_text, p) for p in phrases) if require_all \
           else any(_phrase_in_text(full_text, p) for p in phrases)

def _json_error(message: str, code: int = 500):
    """Return JSON error with a compact traceback string."""
    tb = traceback.format_exc(limit=3)
    resp = jsonify({"success": False, "error": f"{message}", "trace": tb})
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

        filename = secure_filename(f.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_{filename}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        f.save(path)

        # Read Excel
        try:
            if filename.lower().endswith(".xlsx"):
                df = pd.read_excel(path, engine="openpyxl")
            else:
                df = pd.read_excel(path, engine="xlrd")  # requires xlrd==1.2 for .xls
        except Exception:
            df = pd.read_excel(path)

        df = df.fillna("")
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].apply(_clean_text)

        processed_data["current_file"] = path
        processed_data["original_data"] = df

        return jsonify({"success": True, "filename": filename, "rows": int(len(df))})
    except Exception as e:
        return _json_error(f"Upload failed: {e}")

@app.route("/process", methods=["POST"])
def process_file():
    try:
        data = request.get_json(silent=True) or {}
        phrases_in = data.get("additional_keywords", [])
        require_all = bool(data.get("require_all", False))

        if "original_data" not in processed_data:
            return jsonify({"success": False, "error": "No file uploaded"})

        # Normalize phrases (this is the only source of truth)
        phrases: list[str] = []
        for p in phrases_in:
            if isinstance(p, str) and p.strip():
                phrases.append(p.lower().strip())

        df: pd.DataFrame = processed_data["original_data"]  # type: ignore
        matches: list[dict] = []

        if phrases:  # only search when user has active terms
            for _, row in df.iterrows():
                full_text = " ".join(str(row.get(c, "")) for c in df.columns).lower()
                if _matches_by_mode(full_text, phrases, require_all):
                    rd = {k: _clean_text(v) if pd.notna(v) else "" for k, v in row.to_dict().items()}
                    rd["_match_reason"] = "Keyword Match (ALL)" if require_all else "Keyword Match"
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
        return _json_error(f"Processing failed: {e}")

@app.route("/download")
def download_results():
    """Return XLSX in-memory."""
    try:
        if "filtered_data" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})
        rows = [{k: v for k, v in row.items() if k != "_match_reason"} for row in processed_data["filtered_data"]]
        df = pd.DataFrame(rows)
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="results")
        output.seek(0)
        out_name = f"EMAILSIM_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(output, as_attachment=True, download_name=out_name)
    except Exception as e:
        return _json_error(f"Download failed: {e}")

@app.route("/download_csv")
def download_csv():
    """Optional CSV download."""
    try:
        if "filtered_data" not in processed_data:
            return jsonify({"success": False, "error": "No results to download"})
        rows = [{k: v for k, v in row.items() if k != "_match_reason"} for row in processed_data["filtered_data"]]
        df = pd.DataFrame(rows)
        from io import StringIO, BytesIO
        csv_buf = StringIO(); df.to_csv(csv_buf, index=False)
        byte_buf = BytesIO(csv_buf.getvalue().encode("utf-8"))
        out_name = f"EMAILSIM_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        return send_file(byte_buf, as_attachment=True, download_name=out_name, mimetype="text/csv")
    except Exception as e:
        return _json_error(f"CSV export failed: {e}")

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
    import socket
    port = 5000
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", port)) == 0:
            print(f"⚠️  Port {port} in use. Trying 5001…")  # noqa: T201
            port = 5001
    url = f"http://127.0.0.1:{port}/"
    print(f"🌐 Starting web server on: {url}")  # noqa: T201
    t = threading.Thread(target=_open_browser_delayed, args=(url,), daemon=True); t.start()
    try:
        app.run(debug=True, host="127.0.0.1", port=port, use_reloader=False)
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")  # noqa: T201

if __name__ == "__main__":
    main()

# expose the Flask app for gunicorn
application = app
