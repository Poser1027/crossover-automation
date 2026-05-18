"""
FastAPI wrapper for the Hengli Orbital Motor Crossover tool.

Run locally:
    pip install fastapi uvicorn openpyxl
    uvicorn api:app --reload --host 127.0.0.1 --port 8000

Then open http://127.0.0.1:8000 in your browser.

For LAN deployment:
    uvicorn api:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations
from dataclasses import asdict
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from crossover_competitor import crossover_competitor, CrossoverResult
from spec_matching import MatchResult


# ─────────────────────────────────────────────────────────────────────────────
# Serialization helpers
# ─────────────────────────────────────────────────────────────────────────────
def _serialize_match(m: MatchResult) -> dict[str, Any]:
    """Flatten a MatchResult for JSON consumption by the UI."""
    return {
        "part_number": m.suggested_part_number,
        "score": m.score,
        "fit_rating": m.fit_rating,
        "notes": m.notes,
        "caveats": m.caveats,
        "model": {
            "series": m.model.get("series"),
            "type": m.model.get("type"),
            "distribution": m.model.get("distribution"),
            "displacement_cc": m.model.get("displacement_cc"),
            "max_speed_cont": m.model.get("max_speed_cont"),
            "max_speed_inter": m.model.get("max_speed_inter"),
            "max_torque_cont": m.model.get("max_torque_cont"),
            "max_torque_inter": m.model.get("max_torque_inter"),
            "max_dp_cont": m.model.get("max_dp_cont"),
            "max_dp_inter": m.model.get("max_dp_inter"),
            "max_dp_peak": m.model.get("max_dp_peak"),
            "max_flow_cont": m.model.get("max_flow_cont"),
            "max_flow_inter": m.model.get("max_flow_inter"),
            "weight_standard": m.model.get("weight_standard"),
            "weight_bearingless": m.model.get("weight_bearingless"),
        },
        "chosen_codes": m.chosen_codes,
    }


def _serialize_result(r: CrossoverResult) -> dict[str, Any]:
    decoded = None
    if r.decoded:
        decoded = {
            "raw_input": r.decoded.raw_input,
            "normalised": r.decoded.normalised,
            "is_valid": r.decoded.is_valid,
            "brand": r.decoded.brand,
            "series": r.decoded.series,
            "specs": r.decoded.specs_for_crossover,
            "unknown_codes": r.decoded.unknown_codes,
            "warnings": r.decoded.warnings,
        }
    return {
        "raw_input": r.raw_input,
        "is_valid": r.is_valid,
        "decoded": decoded,
        "competitor_brand": r.competitor_brand,
        "competitor_series": r.competitor_series,
        "mapped_hengli_series": r.mapped_hengli_series,
        "mapping_rationale": r.mapping_rationale,
        "primary": [_serialize_match(m) for m in r.primary],
        "fallback": [_serialize_match(m) for m in r.fallback],
        "warnings": r.warnings,
    }


# ─────────────────────────────────────────────────────────────────────────────
# App
# ─────────────────────────────────────────────────────────────────────────────
app = FastAPI(title="Hengli Crossover", version="1.0")

# Permissive CORS so the same machine on a LAN can hit the API from anywhere.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class CrossoverRequest(BaseModel):
    code: str


@app.post("/api/crossover")
def crossover(req: CrossoverRequest) -> dict[str, Any]:
    code = (req.code or "").strip()
    if not code:
        raise HTTPException(status_code=400, detail="Empty code")
    try:
        result = crossover_competitor(code)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Crossover failed: {e}")
    return _serialize_result(result)


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


# Serve the static UI at the root.
STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(STATIC_DIR / "index.html")
