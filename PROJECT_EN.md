# Hengli Orbital Motor Crossover Tool

## ⚡ Collaboration Rules (Read First)

**All responses must prioritise token efficiency:**
- Keep answers concise — no filler or unnecessary explanation
- For code changes, only show the **changed block**, never the whole file
- When asking the user to verify code, **paste the full block for comparison** to avoid partial updates being missed
- Skip reasoning — give conclusions and actions directly
- When multiple changes are needed, deliver all at once in a single response

---

## Background

Parses competitor hydraulic motor model codes and matches them to the closest Hengli equivalent. The user inputs a competitor model code; the system decodes the spec parameters and searches the Hengli product database for the best match.

Available as both a Python CLI and a local web UI (see Web UI section below).

---

## Tech Stack

- Language: Python 3.9+
- Core dependency: openpyxl
- Web UI dependencies: fastapi, uvicorn
- Run mode: CLI, programmatic import, or browser-based UI

---

## File Structure

| File | Description |
|------|-------------|
| `crossover_competitor.py` | Main entry point: competitor code → Hengli match output |
| `decode_competitor.py` | Competitor code decoder (brand, series, specs) |
| `spec_matching.py` | Hengli database loader and scoring/matching logic |
| `decode_part_number.py` | Hengli part-number decoder (reverse direction) |
| `Hengli_Orbital_Motor_Master.xlsx` | Hengli product database |
| `competitor_code_extractor.xlsx` | Competitor code decode rule table |
| `api.py` | FastAPI wrapper exposing the engine over HTTP |
| `static/index.html` | Web UI (single page, vanilla JS) |
| `run.bat` / `run.sh` | One-click launchers for Windows / macOS / Linux |

---

## Series Mapping

```python
SERIES_MAPPING = {
    ("Char-Lynn (Eaton)", "2000 Series") → Hengli HSP/HSD
    ("Char-Lynn (Eaton)", "T Series")    → Hengli HRD
}
```

To add a new brand/series, add an entry to `SERIES_MAPPING` in `crossover_competitor.py`.

---

## Key Parameters & Thresholds

| Parameter | Value | Description |
|-----------|-------|-------------|
| `MIN_SCORE` | 50.0 | Candidates below this score are discarded |
| Displacement lock tolerance | ±10% | All three candidates must fall within this range of the anchor |
| Shaft diameter exact match | ±0.1 mm | Preferred |
| Shaft diameter loose match | ±1.0 mm | Fallback only if no exact match |
| `WEAK_PRIMARY_THRESHOLD` | 50.0 | Primary score below this triggers fallback search |
| `FALLBACK_TOP_N` | 3 | Max fallback candidates |
| `top_n` (default) | 3 | Max primary candidates |

---

## Matching Logic

### Displacement Scoring (`_score_model`)

| Deviation | Score |
|-----------|-------|
| ≤ 5% | 50 (full) |
| 5–15% | 50 → 20 linear |
| 15–30% | 20 → 2 linear |
| > 30% | ~0 |

### Three-Candidate Selection (`find_matches`)

1. Discard all candidates with score < 50
2. Use the highest-scoring model's displacement as the **anchor**
3. Only consider candidates within ±10% of the anchor displacement
4. Keep only the highest-scoring model per unique displacement value → **three candidates with different displacements**
5. Return fewer than 3 if not enough qualify — never pad with weak matches

### Shaft Diameter Selection (`_pick_shaft`)

- First: exact match ±0.1 mm (handles float rounding)
- Fallback: loose match ±1.0 mm
- Last resort: series default code

### Shaft Description Regex

Handles two formats:
- With Ø prefix: `Ø32 straight...`
- Without prefix: `31.75 (1.250) dia straight shaft...`

Regex: `(?:ø\s*)?([0-9]+(?:\.[0-9]+)?)\s*(?:mm|\(|straight|spline|taper)`

### Shaft Diameter Decoding (`decode_competitor.py`)

Competitor descriptions may put mm first (`31.75 (1.250)`) or inside parentheses (`1 inch (25.4 mm)`):
- Prefer `(number mm)` pattern
- Fall back to "number before opening parenthesis"

---

## Default Fill Values

Fields left as `None` by the decoder are auto-filled before matching:

| Field | Default |
|-------|---------|
| `special` | `"A"` |
| `paint` | `"N"` |
| `rotation` | `"CW"` |

---

## Usage

### CLI

```bash
python crossover_competitor.py "M02 02 049 AC 02 AA 01 0 00 1 0 00 00 00 AA AA F"
python crossover_competitor.py M02049AC02AA0100010000000AAAAF
python crossover_competitor.py   # interactive prompt
```

### Programmatic

```python
from crossover_competitor import crossover_competitor, print_crossover
result = crossover_competitor("M02049AC02AA0100010000000AAAAF")
print_crossover(result)
```

### Web UI

**Windows:** double-click `run.bat`
**macOS / Linux:** `./run.sh`

Opens the UI at `http://127.0.0.1:8000`. First run auto-installs `fastapi`, `uvicorn`, `openpyxl`.

For LAN sharing: `uvicorn api:app --host 0.0.0.0 --port 8000`

---

## Web UI

### Architecture

```
Browser (static/index.html)
  ↓ POST /api/crossover {"code": "..."}
FastAPI (api.py)
  ↓ calls crossover_competitor()
Existing engine (unchanged)
```

Stateless. Single HTML page, no build step. Zero changes to the existing engine.

### API Endpoints

| Endpoint | Method | Body | Returns |
|----------|--------|------|---------|
| `/api/crossover` | POST | `{"code": "..."}` | Full crossover JSON |
| `/api/health` | GET | — | `{"status": "ok"}` |
| `/` | GET | — | Serves the UI |

### JSON Response Shape

```
{
  raw_input, is_valid,
  decoded: { brand, series, normalised, specs, warnings },
  competitor_brand, competitor_series,
  mapped_hengli_series, mapping_rationale,
  primary:  [{ part_number, score, fit_rating, model{...}, chosen_codes{...}, caveats }],
  fallback: [...same shape...],
  warnings: []
}
```

### UI Features

- Single input box, Enter or button to submit
- Decoded specs panel (displacement, pressure, speed, flow, torque, mount, shaft, port, rotation)
- Series mapping line with Hengli rationale
- Primary candidate cards with color-coded score badge (green ≥70, yellow 50–69, red <50), fit rating, full ratings table (cont/inter/peak), chosen codes, caveats
- Fallback candidate cards (when triggered)
- Warnings panel
- **Copy** button per part number
- **Export CSV** / **Export JSON** for full result
- **Recent Lookups** chips — last 20 codes in `localStorage`, click to re-run

### Deployment Modes

| Mode | Command | Use |
|------|---------|-----|
| Local | `run.bat` or `run.sh` | Single user, default 127.0.0.1 |
| LAN | `uvicorn api:app --host 0.0.0.0 --port 8000` | Internal team access |
| Production | Reverse proxy (nginx) + systemd / Windows service | Permanent install |

---

## Progress

- [x] Competitor model code decoding
- [x] Series mapping (Char-Lynn 2000 → HSP/HSD, T → HRD)
- [x] Primary match + fallback match
- [x] Shaft diameter exact match (fixed 1 inch → 25mm mismatch)
- [x] Shaft diameter decoding (fixed mm/inch order ambiguity)
- [x] Shaft description regex supports both Ø-prefixed and bare formats
- [x] Displacement penalty weight increased
- [x] Three-candidate displacement diversity (no more identical-displacement triplicates)
- [x] Candidates with score < 50 discarded
- [x] FastAPI wrapper (`api.py`) with JSON serialization of `CrossoverResult`
- [x] Web UI (single page, vanilla JS) — lookup, history, copy, CSV/JSON export
- [x] One-click launchers (`run.bat`, `run.sh`)
- [ ] Batch lookup (paste multiple codes at once)
- [ ] Additional competitor brands (Parker, Danfoss, Bosch Rexroth)

---

## Known Issues & Feedback

(None)

---

## Decision Log

- `SERIES_MAPPING` dict over hardcoded if/else — easier to extend
- Fallback logic: primary-first, expand only if weak — avoids noise from full-db search
- Default values (special/paint/rotation) filled at crossover layer, not in decoder — keeps decode output clean
- Shaft diameter: exact-first, loose-fallback — prevents 25.4mm matching to 25mm
- Three candidates must differ in displacement — prevents redundant recommendations
- Score threshold 50 = lower bound of "Functional" fit rating — below this, manual review is more appropriate than auto-recommendation
- **UI stack: FastAPI + vanilla HTML** — over Streamlit/Tkinter. No build step, no per-machine install for end users, runs on a single laptop or LAN with the same code
- **UI is a thin wrapper** — `api.py` calls `crossover_competitor()` unchanged. All business logic stays in the engine, the UI just renders results
- **localStorage history, not server-side** — keeps the app stateless, no DB to maintain, history is per-user/per-browser
- **CSV + JSON export** — CSV for sales/spreadsheet workflows, JSON for downstream integration or debugging

---

## Common Pitfalls When Editing Code

1. **Clear cache after changes**: delete `__pycache__/` or Python will still run the old `.pyc`
2. **Verify all changed blocks**: when asking the user to check, paste the full block — partial updates are easy to miss
3. **Never assume description format**: Hengli and competitor descriptions differ (Ø prefix or not, mm/inch order) — use permissive regex + downstream filtering
4. **UI changes don't need a server restart for HTML/JS** — uvicorn reloads `api.py` automatically if launched with `--reload`, but `static/index.html` is always served fresh on hard refresh (Ctrl+Shift+R)
5. **JSON shape is a contract** — if you add a field to `CrossoverResult` or `MatchResult`, also update `_serialize_match` / `_serialize_result` in `api.py`, or it won't reach the UI
