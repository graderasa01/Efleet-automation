from __future__ import annotations

import shutil
import threading
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .config_loader import load_config
from .control_hub import ControlHub
from .efleet_worker import EfleetWorker
from .excel_manager import ExcelManager
from .query_engine import QUERY_MASTER, query_options_for_ui

app = FastAPI(title="E-fleet Codespaces Control Panel")
app.mount("/static", StaticFiles(directory="static"), name="static")

CFG = load_config()
HUB = ControlHub()
EXCEL = ExcelManager(
    CFG.get("excel_file", "data/VehicleDetails.xlsx"),
    CFG.get("sheet_name", "DATA"),
    CFG.get("search_column_name", "HireSlipNo"),
    CFG.get("search_column_index_1_based", 2),
)
WORKER_THREAD: Optional[threading.Thread] = None


def worker_is_alive() -> bool:
    return WORKER_THREAD is not None and WORKER_THREAD.is_alive()


@app.get("/", response_class=HTMLResponse)
def home() -> str:
    options = query_options_for_ui()
    option_html = "\n".join(
        f'<option value="{o["column"]}|{o["query_no"]}">{o["key"]} → {o["column"]} [{o["query_no"]}] — {o["text"]}</option>'
        for o in options
    )
    return f"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>E-fleet Control Panel</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 0; background: #f5f7fb; color: #111827; }}
    .wrap {{ max-width: 1200px; margin: 0 auto; padding: 18px; }}
    .card {{ background: white; border: 1px solid #e5e7eb; border-radius: 14px; padding: 16px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }}
    h1 {{ margin: 0 0 8px; font-size: 26px; }}
    h2 {{ margin: 0 0 12px; font-size: 18px; }}
    .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .buttons {{ display: flex; flex-wrap: wrap; gap: 8px; }}
    button {{ border: 0; border-radius: 10px; padding: 10px 14px; font-weight: 700; cursor: pointer; }}
    .green {{ background: #16a34a; color: white; }}
    .blue {{ background: #2563eb; color: white; }}
    .yellow {{ background: #f59e0b; color: white; }}
    .red {{ background: #dc2626; color: white; }}
    .gray {{ background: #374151; color: white; }}
    .light {{ background: #e5e7eb; color: #111827; }}
    table {{ border-collapse: collapse; width: 100%; }}
    td {{ padding: 6px 8px; border-bottom: 1px solid #f1f5f9; }}
    td:first-child {{ color: #64748b; width: 180px; }}
    select, input[type=text] {{ width: 100%; padding: 10px; border: 1px solid #d1d5db; border-radius: 10px; }}
    .imgbox {{ min-height: 260px; display:flex; align-items:center; justify-content:center; background:#111827; border-radius: 12px; overflow:hidden; }}
    .imgbox img {{ max-width: 100%; max-height: 520px; object-fit: contain; }}
    .small {{ font-size: 12px; color: #6b7280; }}
    .pill {{ display: inline-block; padding: 4px 8px; border-radius: 999px; background: #eef2ff; color:#3730a3; font-weight:700; }}
  </style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>🚚 E-fleet Automation Control</h1>
    <div class="small">Codespaces me ye panel port <b>8000</b> par chalega. Excel file server-side update hogi.</div>
  </div>

  <div class="grid">
    <div class="card">
      <h2>Bot Status</h2>
      <table id="statusTable"></table>
    </div>
    <div class="card">
      <h2>Main Buttons</h2>
      <div class="buttons">
        <button class="green" onclick="cmd('START')">▶ Start / Resume</button>
        <button class="yellow" onclick="cmd('PAUSE')">⏸ Pause</button>
        <button class="blue" onclick="cmd('SAVE_NEXT')">💾 Save & Next</button>
        <button class="light" onclick="cmd('SKIP')">⏭ Skip HM</button>
        <button class="light" onclick="cmd('CLOSE_IMAGE')">🖼 Close Image</button>
        <button class="light" onclick="cmd('STOP_IMAGES')">🛑 Stop Images</button>
        <button class="gray" onclick="cmd('REOPEN_HM')">🔁 Reopen HM</button>
        <button class="red" onclick="cmd('STOP')">⛔ Stop Bot</button>
      </div>
      <hr>
      <div class="buttons">
        <button class="green" onclick="startBot()">🚀 Launch Bot Worker</button>
        <button class="light" onclick="window.location='/download-excel'">⬇ Download Excel</button>
      </div>
    </div>
  </div>

  <div class="card">
    <h2>Apply Query to Current HM</h2>
    <div style="display:grid; grid-template-columns: 2fr 1fr; gap:10px; align-items:end;">
      <div>
        <label class="small">Preset query</label>
        <select id="querySelect">{option_html}</select>
      </div>
      <button class="blue" onclick="applyQuery()">✍ Apply Query</button>
    </div>
    <div style="margin-top:10px; display:grid; grid-template-columns: 1fr 2fr 1fr; gap:10px; align-items:end;">
      <div><label class="small">Custom column</label><input id="customColumn" type="text" placeholder="Insurance"></div>
      <div><label class="small">Custom text</label><input id="customText" type="text" placeholder="Type custom remark"></div>
      <button class="gray" onclick="applyCustom()">✍ Apply Custom</button>
    </div>
  </div>

  <div class="grid">
    <div class="card">
      <h2>Current Image Preview</h2>
      <div class="imgbox"><img id="currentImage" src="" onerror="this.style.display='none'" /></div>
      <div class="small">Image screenshot auto refresh hota hai. Agar blank ho to image nahi mili ya screenshot fail hua.</div>
    </div>
    <div class="card">
      <h2>Upload / Replace Excel</h2>
      <form action="/upload-excel" method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx" required>
        <button class="blue" type="submit">Upload Excel</button>
      </form>
      <p class="small">Upload ke baad Launch Bot Worker dabao. File path config me <code>{CFG.get('excel_file')}</code> hai.</p>
      <hr>
      <h2>Important</h2>
      <p class="small">Credentials ko repo me commit mat karna. Codespaces terminal me environment variables set karo.</p>
    </div>
  </div>
</div>
<script>
async function status() {{
  const res = await fetch('/api/status');
  const s = await res.json();
  const keys = ['running','paused','stopped','current_index','total_rows','current_excel_row','current_hm','current_image_index','total_images','message','last_error','updated_at'];
  document.getElementById('statusTable').innerHTML = keys.map(k => `<tr><td>${{k}}</td><td><span class="pill">${{s[k] ?? ''}}</span></td></tr>`).join('');
  const img = document.getElementById('currentImage');
  if (s.current_image_path) {{ img.style.display='block'; img.src = s.current_image_path + '?t=' + Date.now(); }}
}}
async function cmd(name, payload={{}}) {{
  await fetch('/api/command', {{method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{name, payload}})}});
  await status();
}}
async function startBot() {{
  const res = await fetch('/api/start-worker', {{method:'POST'}});
  const txt = await res.text();
  console.log(txt);
  await status();
}}
async function applyQuery() {{
  const [column, query_no] = document.getElementById('querySelect').value.split('|');
  await cmd('APPLY_QUERY', {{column, query_no}});
}}
async function applyCustom() {{
  const column = document.getElementById('customColumn').value;
  const custom_text = document.getElementById('customText').value;
  await cmd('APPLY_QUERY', {{column, query_no:'', custom_text}});
}}
setInterval(status, 1000);
status();
</script>
</body>
</html>
"""


@app.get("/api/status")
def api_status():
    return HUB.get_status()


@app.post("/api/start-worker")
def start_worker():
    global WORKER_THREAD, EXCEL
    if worker_is_alive():
        return JSONResponse({"ok": True, "message": "Worker already running"})
    # Recreate Excel manager in case user uploaded/replaced workbook.
    EXCEL = ExcelManager(
        CFG.get("excel_file", "data/VehicleDetails.xlsx"),
        CFG.get("sheet_name", "DATA"),
        CFG.get("search_column_name", "HireSlipNo"),
        CFG.get("search_column_index_1_based", 2),
    )
    worker = EfleetWorker(CFG, HUB, EXCEL)
    WORKER_THREAD = threading.Thread(target=worker.run, daemon=True)
    WORKER_THREAD.start()
    return {"ok": True, "message": "Worker launched"}


@app.post("/api/command")
async def api_command(body: dict):
    name = str(body.get("name", "")).upper().strip()
    payload = body.get("payload") or {}
    if not name:
        raise HTTPException(status_code=400, detail="Missing command name")

    if name == "APPLY_QUERY":
        st = HUB.get_status()
        excel_row = st.get("current_excel_row")
        if not excel_row:
            raise HTTPException(status_code=400, detail="No current Excel row active")
        column = str(payload.get("column", "")).strip()
        query_no = str(payload.get("query_no", "")).strip()
        custom_text = str(payload.get("custom_text", "")).strip()
        if not column:
            raise HTTPException(status_code=400, detail="Missing column")
        try:
            audit = EXCEL.apply_query(int(excel_row), column, query_no, custom_text)
            HUB.update_status(message=f"Applied query in {column}. Audit: {audit[:120]}")
        except Exception as e:
            HUB.update_status(last_error=str(e), message=f"Apply query failed: {str(e)[:120]}")
            raise HTTPException(status_code=400, detail=str(e))
    else:
        HUB.push_command(name, payload)
    return {"ok": True, "status": HUB.get_status()}


@app.post("/upload-excel")
def upload_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx file supported")
    path = Path(CFG.get("excel_file", "data/VehicleDetails.xlsx"))
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    HUB.update_status(message=f"Uploaded Excel: {path}")
    return HTMLResponse("<p>Excel uploaded. <a href='/'>Back to panel</a></p>")


@app.get("/download-excel")
def download_excel():
    path = Path(CFG.get("excel_file", "data/VehicleDetails.xlsx"))
    if not path.exists():
        raise HTTPException(status_code=404, detail="Excel file not found")
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
