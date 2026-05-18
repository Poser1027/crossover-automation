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

---

## Tech Stack

- Language: Python 3.9+
- Dependency: openpyxl
- Run mode: CLI or programmatic import

---

## File Structure

| File | Description |
|------|-------------|
| `crossover_competitor.py` | Main entry point: competitor code → Hengli match output |
| `decode_competitor.py` | Competitor code decoder (brand, series, specs) |
| `spec_matching.py` | Hengli database loader and scoring/matching logic |
| `Hengli_Orbital_Motor_Master.xlsx` | Hengli product database |
| `competitor_code_extractor.xlsx` | Competitor code decode rule table |

---

## Series Mapping

```python
SERIES_MAPPING = {
    ("Char-Lynn (Eaton)", "2000 Series") → Hengli HSP
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

```bash
# CLI
python crossover_competitor.py "M02 02 049 AC 02 AA 01 0 00 1 0 00 00 00 AA AA F"
python crossover_competitor.py M02049AC02AA0100010000000AAAAF
python crossover_competitor.py   # interactive prompt

# Programmatic
from crossover_competitor import crossover_competitor, print_crossover
result = crossover_competitor("M02049AC02AA0100010000000AAAAF")
print_crossover(result)
```

---

## Progress

- [x] Competitor model code decoding
- [x] Series mapping (Char-Lynn 2000 → HSP, T → HRD)
- [x] Primary match + fallback match
- [x] Shaft diameter exact match (fixed 1 inch → 25mm mismatch)
- [x] Shaft diameter decoding (fixed mm/inch order ambiguity)
- [x] Shaft description regex supports both Ø-prefixed and bare formats
- [x] Displacement penalty weight increased
- [x] Three-candidate displacement diversity (no more identical-displacement triplicates)
- [x] Candidates with score < 50 discarded
- [ ] (to be added)

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

---

## Common Pitfalls When Editing Code

1. **Clear cache after changes**: delete `__pycache__/` or Python will still run the old `.pyc`
2. **Verify all changed blocks**: when asking the user to check, paste the full block — partial updates are easy to miss
3. **Never assume description format**: Hengli and competitor descriptions differ (Ø prefix or not, mm/inch order) — use permissive regex + downstream filtering
