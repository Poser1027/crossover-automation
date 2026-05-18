# Hengli Crossover — Web UI

Minimal local web app wrapping the existing Python crossover engine. No changes to your existing `.py` files or `.xlsx` files — this is a thin layer on top.

## File Layout

Drop these alongside your existing files. Final folder should look like:

```
your_folder/
├── api.py                              ← NEW (FastAPI wrapper)
├── run.bat                             ← NEW (Windows launcher)
├── run.sh                              ← NEW (macOS / Linux launcher)
├── static/
│   └── index.html                      ← NEW (UI)
│
├── crossover_competitor.py             ← existing
├── decode_competitor.py                ← existing
├── spec_matching.py                    ← existing
├── decode_part_number.py               ← existing
├── Hengli_Orbital_Motor_Master.xlsx    ← existing
└── competitor_code_extractor.xlsx      ← existing
```

## First-Time Setup

**Windows**: double-click `run.bat`. It installs dependencies the first time, then opens the UI in your browser.

**macOS / Linux**:
```bash
chmod +x run.sh
./run.sh
```

If you prefer manual setup:
```bash
pip install fastapi uvicorn openpyxl
uvicorn api:app --host 127.0.0.1 --port 8000
```

Then open <http://127.0.0.1:8000>

## Usage

1. Paste a competitor model code (e.g. `M02049AC02AA0100010000000AAAAF`)
2. Click **Cross Over**
3. Review decoded specs → series mapping → primary candidates → alternatives
4. Click **Copy** on any part number, or **Export CSV / JSON** for the whole result
5. **Recent Lookups** chips let you re-run any previous code in one click

## Deploying to an Internal Server (later)

Run with host `0.0.0.0` to expose on your LAN:

```bash
uvicorn api:app --host 0.0.0.0 --port 8000
```

Then teammates can open `http://<your-machine-ip>:8000`.

For a more permanent install, run behind a reverse proxy (nginx) or use a process manager (systemd, Windows service, `pm2`). The API itself is a stateless FastAPI app — standard deployment patterns apply.

## API

`POST /api/crossover`
```json
{ "code": "M02049AC02AA0100010000000AAAAF" }
```
Returns the full crossover result as JSON (decoded specs, mapping, primary, fallback, warnings).

`GET /api/health` → `{"status": "ok"}`
