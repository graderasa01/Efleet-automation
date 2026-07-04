# E-fleet Codespaces Automation Bot

This is a **GitHub Codespaces-friendly** E-fleet automation project.

It gives you:

- Playwright Python automation for E-fleet HM/HireSlip search
- Browser-based control panel with buttons
- Excel `.xlsx` read/write using `openpyxl`
- Query apply system for audit columns
- Current image screenshot preview in the control panel
- Save & Next, Pause, Resume, Skip, Close Image, Stop Images, Reopen HM, Stop Bot

> Important: Codespaces does not run local Microsoft Excel UI. So live Excel VBA buttons will not work inside Codespaces. This project replaces Excel buttons with a web control panel at port `8000`. Excel is still connected as the data/output file.

---

## 1. Folder structure

```text
.
├─ main.py
├─ requirements.txt
├─ config.example.json
├─ .env.example
├─ .devcontainer/devcontainer.json
├─ efleet_bot/
│  ├─ config_loader.py
│  ├─ control_hub.py
│  ├─ efleet_worker.py
│  ├─ excel_manager.py
│  ├─ query_engine.py
│  └─ web_app.py
├─ data/
│  ├─ sample_vehicle_details.xlsx
│  └─ VehicleDetails.xlsx
└─ static/
   └─ current_image.png  (created during run)
```

---

## 2. Codespaces setup

1. Push this folder to GitHub repo.
2. Open repo → **Code** → **Codespaces** → **Create codespace**.
3. Wait for setup. The devcontainer will run:

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

If setup did not run automatically, run manually:

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

---

## 3. Set credentials safely

Do **not** commit real username/password to GitHub.

In Codespaces terminal:

```bash
export EFLEET_URL="http://your-efleet-url/VehiclePlacement/PlacementDashboard"
export EFLEET_USERNAME="your_username"
export EFLEET_PASSWORD="your_password"
export EFLEET_CONFIG="config.example.json"
```

For persistent Codespaces secrets, use GitHub Codespaces secrets from GitHub settings.

---

## 4. Add your Excel file

Put your workbook here:

```text
data/VehicleDetails.xlsx
```

Or open the control panel and upload Excel.

The workbook should have a `DATA` sheet. If it does not, the bot uses the active sheet.

Default search field:

```json
"search_column_name": "HireSlipNo",
"search_column_index_1_based": 2
```

If duplicate `HireSlipNo` columns exist, the bot prefers exact header match first; otherwise it falls back to first header containing `hireslip`, then numeric index.

---

## 5. Run control panel

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

Codespaces will auto-forward port `8000`. Open the forwarded URL.

---

## 6. Button workflow

1. Open control panel.
2. Press **Launch Bot Worker**.
3. Press **Start / Resume**.
4. Bot logs in and starts HM processing.
5. During image/document review, use:
   - Pause
   - Resume
   - Close Image
   - Stop Images
   - Apply Query
   - Save & Next
   - Skip HM
   - Reopen HM
   - Stop Bot
6. Download updated Excel from **Download Excel**.

---

## 7. Query system

Preset query examples:

| Key | Column | Query |
|---|---|---|
| e | E- Way bill | E-Way bill is not Available in efleet |
| t | Tax invoice | Tax invoice is not found in efleet |
| w | weight slip | weight slip is not found |
| v | Vehicle photo | Vehicle photo is not uploaded |
| i | Insurance | Insurance Expired |
| n | Fitness | Fitness Expired |
| u | PUC | PUC Expired |
| d | DL | DL is not found |

When you apply a query, the bot writes it into the current row and rebuilds `Audit Queries` automatically.

---

## 8. Important limitations

### Codespaces limitation

Codespaces is a remote Linux environment. It cannot control your local Microsoft Excel buttons. This is why this project uses a web panel as the button/control layer.

### Browser visibility

This config uses:

```json
"headless": true
```

That is best for Codespaces. If you need a visible remote browser, you need Xvfb/noVNC setup. For normal office work, the web image screenshot preview is simpler.

---

## 9. First test recommendation

Before using full file, test with 1–2 rows.

You can temporarily set:

```json
"max_images": 3
```

Then run and verify:

- HM search works
- eye/view opens
- image section opens
- query writes to Excel
- Save & Next works
- output Excel downloads correctly

---

## 10. If selector fails

Open `config.example.json` and update these selectors:

```json
"search_box"
"eye_selectors"
"step2_selectors"
"image_step_selectors"
"image_link_selectors"
```

The bot has fallback selectors, but E-fleet page layout changes can still require selector adjustment.
