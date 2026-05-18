"""
Competitor → Hengli Crossover
==============================

Take a competitor's orbital motor model code (or a pre-decoded result from
decode_competitor.py) and return the best-matching Hengli orbital motor
candidates, using the series mapping defined at the top of this file.

Currently mapped series:
    Char-Lynn (Eaton) 2000 Series  →  Hengli HSP
    Char-Lynn (Eaton) T Series     →  Hengli HRD

Behavior:
  - PRIMARY: search only the mapped Hengli series for the best matches.
  - FALLBACK: if no primary candidates exist (or the top primary score is < 50),
    also search other Hengli series and offer them as alternatives.

USAGE
-----
    # CLI — accepts a competitor model code as the argument
    python crossover_competitor.py "M02 02 049 AC 02 AA 01 0 00 1 0 00 00 00 AA AA F"
    python crossover_competitor.py M02049AC02AA0100010000000AAAAF
    python crossover_competitor.py                  # interactive prompt

    # Programmatic — accepts either a raw code string OR a CompetitorDecodeResult
    from crossover_competitor import crossover_competitor, print_crossover

    result = crossover_competitor("M02049AC02AA0100010000000AAAAF")
    print_crossover(result)

REQUIREMENTS
------------
- Python 3.9+, openpyxl
- spec_matching.py (with the series parameter on find_matches)
- decode_competitor.py
- Hengli_Orbital_Motor_Master.xlsx
- competitor_code_extractor.xlsx
- All four files in the same folder as this script
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union
import sys

from spec_matching import MotorDatabase, find_matches, MatchResult, check_input_sufficiency
from decode_competitor import decode_competitor_code, CompetitorDecodeResult


# ─────────────────────────────────────────────────────────────────────────────
# Series mapping — competitor (brand, series) → Hengli series
# ─────────────────────────────────────────────────────────────────────────────
SERIES_MAPPING: dict[tuple[str, str], dict] = {
    ("Char-Lynn (Eaton)", "2000 Series"): {
        "hengli_series": ["HSP", "HSD"],
        "rationale": (
            "HSP and HSD are the disc-valve, premium-pressure families "
            "(225 bar continuous) with wheel mount and high-pressure shaft "
            "seal options — the closest structural match for the Char-Lynn "
            "2000 Series."
        ),
    },
    ("Char-Lynn (Eaton)", "T Series"): {
        "hengli_series": "HRD",
        "rationale": (
            "HRD is the spool-valve, light/medium-duty family — best match for "
            "the Char-Lynn T Series, which uses a similar duty profile."
        ),
    },
}

# Fallback trigger thresholds
WEAK_PRIMARY_THRESHOLD = 50.0   # if top primary score < this, also show alternatives
FALLBACK_TOP_N = 3              # how many alternative-series candidates to show


# ─────────────────────────────────────────────────────────────────────────────
# Result container
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class CrossoverResult:
    """Full crossover outcome."""
    raw_input: str
    decoded: Optional[CompetitorDecodeResult]
    competitor_brand: Optional[str]
    competitor_series: Optional[str]
    mapped_hengli_series: Optional[str]
    mapping_rationale: str = ""
    primary: list[MatchResult] = field(default_factory=list)
    fallback: list[MatchResult] = field(default_factory=list)
    is_valid: bool = False
    warnings: list[str] = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────
def crossover_competitor(
    input_data: Union[str, CompetitorDecodeResult],
    hengli_workbook: Optional[str] = None,
    competitor_workbook: Optional[str] = None,
    top_n: int = 3,
) -> CrossoverResult:
    """
    Convert a competitor model code into Hengli equivalent recommendations.

    Args:
        input_data: either a raw competitor model code string,
                    or a CompetitorDecodeResult from decode_competitor.py
        hengli_workbook: path to Hengli_Orbital_Motor_Master.xlsx
                         (defaults to script folder)
        competitor_workbook: path to competitor_code_extractor.xlsx
                             (defaults to script folder; only used when input is a string)
        top_n: number of primary Hengli candidates to return

    Returns:
        CrossoverResult with primary matches, fallback alternatives, and warnings.
    """
    base_dir = Path(__file__).parent
    if hengli_workbook is None:
        hengli_workbook = str(base_dir / "Hengli_Orbital_Motor_Master.xlsx")
    if competitor_workbook is None:
        competitor_workbook = str(base_dir / "competitor_code_extractor.xlsx")

    # Step 1 — Resolve input to a CompetitorDecodeResult
    if isinstance(input_data, CompetitorDecodeResult):
        decoded = input_data
        raw_input = decoded.raw_input
    elif isinstance(input_data, str):
        raw_input = input_data
        decoded = decode_competitor_code(raw_input, competitor_workbook)
    else:
        raise TypeError(
            f"input_data must be a string or CompetitorDecodeResult, "
            f"got {type(input_data).__name__}"
        )

    result = CrossoverResult(
        raw_input=raw_input,
        decoded=decoded,
        competitor_brand=decoded.brand,
        competitor_series=decoded.series,
        mapped_hengli_series=None,
    )

    # Step 2 — Did the decode itself succeed?
    if not decoded.is_valid:
        result.warnings.append(
            "Competitor decode failed — cannot run crossover. "
            "See decode warnings below."
        )
        result.warnings.extend(decoded.warnings)
        return result

    # Step 3 — Look up the mapping
    mapping_key = (decoded.brand, decoded.series)
    mapping = SERIES_MAPPING.get(mapping_key)
    if mapping is None:
        result.warnings.append(
            f"No series mapping defined for {decoded.brand} {decoded.series}. "
            f"Add an entry to SERIES_MAPPING in crossover_competitor.py."
        )
        return result

    mapped_series = mapping["hengli_series"]
    mapped_series_list = [mapped_series] if isinstance(mapped_series, str) else list(mapped_series)
    result.mapped_hengli_series = "/".join(mapped_series_list)
    result.mapping_rationale = mapping.get("rationale", "")

    # Step 4 — Load the Hengli database and check spec sufficiency
    if not Path(hengli_workbook).exists():
        result.warnings.append(f"Hengli workbook not found: {hengli_workbook}")
        return result
    hengli_db = MotorDatabase(hengli_workbook)

    specs = decoded.specs_for_crossover
    # Fill in defaults for fields the decoder leaves as None (find_matches needs them
    # as strings, not None, when building the part number).
    if specs.get("special") is None:
        specs["special"] = "A"
    if specs.get("paint") is None:
        specs["paint"] = "N"
    if specs.get("rotation") is None:
        specs["rotation"] = "CW"

    ok, msg = check_input_sufficiency(specs)
    if not ok:
        result.warnings.append(f"Insufficient specs for matching: {msg}")
        return result

    # Step 5 — Carry over decoder warnings (e.g. ratings missing from code)
    result.warnings.extend(
        w for w in decoded.warnings
        if "not encoded" in w or "inferred" in w
    )

    # Step 6 — PRIMARY search (mapped series only)
    try:
        result.primary = find_matches(specs, hengli_db, top_n=top_n,
                                      series=mapped_series_list)
    except ValueError as e:
        result.warnings.append(f"Primary search failed: {e}")

    # Step 7 — FALLBACK search (only if primary empty or weak)
    needs_fallback = (
        not result.primary
        or (result.primary and result.primary[0].score < WEAK_PRIMARY_THRESHOLD)
    )
    if needs_fallback:
        try:
            all_matches = find_matches(specs, hengli_db, top_n=FALLBACK_TOP_N + 5)
            fallback = [m for m in all_matches
                        if m.model["series"] not in mapped_series_list]
            result.fallback = fallback[:FALLBACK_TOP_N]
        except ValueError:
            pass

        if result.fallback:
            if not result.primary:
                result.warnings.append(
                    f"No {mapped_series} candidates met the requirements. "
                    f"Alternative-series candidates shown below."
                )
            else:
                result.warnings.append(
                    f"Primary {mapped_series} candidates scored below "
                    f"{WEAK_PRIMARY_THRESHOLD:.0f}. Stronger candidates from "
                    f"other Hengli series shown as alternatives."
                )

    result.is_valid = bool(result.primary or result.fallback)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Pretty-printed report
# ─────────────────────────────────────────────────────────────────────────────
def _fmt(v, suffix: str = "") -> str:
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:g}{suffix}"
    return f"{v}{suffix}"


def _print_match_block(m: MatchResult, index: int, label: str) -> None:
    print("─" * 78)
    print(f" {label} #{index}   |   Part No.: {m.suggested_part_number}   |   "
          f"Fit: {m.fit_rating}   |   Score: {m.score}/100")
    print("─" * 78)

    mdl = m.model
    print(f"  Series / Frame:      {mdl['series']} / {mdl['type']}")
    print(f"  Distribution:        {mdl.get('distribution', '—')}")
    print(f"  Displacement:        {_fmt(mdl.get('displacement_cc'))} cc/rev")
    print()
    print(f"  RATINGS                Continuous   Intermittent   Peak")
    print(f"  Max speed (rpm)        {_fmt(mdl.get('max_speed_cont')):>10}   "
          f"{_fmt(mdl.get('max_speed_inter')):>12}   —")
    print(f"  Max torque (N·m)       {_fmt(mdl.get('max_torque_cont')):>10}   "
          f"{_fmt(mdl.get('max_torque_inter')):>12}   —")
    print(f"  Max Δp (bar)           {_fmt(mdl.get('max_dp_cont')):>10}   "
          f"{_fmt(mdl.get('max_dp_inter')):>12}   "
          f"{_fmt(mdl.get('max_dp_peak')):>5}")
    print(f"  Max flow (L/min)       {_fmt(mdl.get('max_flow_cont')):>10}   "
          f"{_fmt(mdl.get('max_flow_inter')):>12}   —")
    print(f"  Weight:                std {_fmt(mdl.get('weight_standard'))} kg, "
          f"bearingless {_fmt(mdl.get('weight_bearingless'))} kg")

    cc = m.chosen_codes
    print()
    print(f"  CHOSEN CODES (auto-picked):")
    print(f"  Mount    {cc['mount_code']:<6}  {cc['mount_desc']}")
    if cc.get("port_code"):
        print(f"  Port     {cc['port_code']:<6}  {cc['port_desc']}")
    print(f"  Shaft    {cc['shaft_code']:<6}  {cc['shaft_desc']}")
    if cc.get("shaft_max_torque"):
        print(f"           (shaft max torque: {cc['shaft_max_torque']} N·m)")
    print(f"  Rotation {cc['rotation_code']}     Paint {cc['paint_code']}     "
          f"Special {cc['special_code']}")

    if m.caveats:
        print()
        print("  Caveats:")
        for c in m.caveats:
            print(f"    ⚠ {c}")
    print()


def print_crossover(result: CrossoverResult) -> None:
    """Pretty-printed competitor → Hengli crossover report."""
    print()
    print("═" * 78)
    print(" COMPETITOR → HENGLI CROSSOVER")
    print("═" * 78)

    print(f"\n  Competitor input:    {result.raw_input}")
    if result.decoded:
        print(f"  Brand:               {result.competitor_brand}")
        print(f"  Series:              {result.competitor_series}")
        print(f"  Normalised code:     {result.decoded.normalised}")

    if result.decoded:
        specs = result.decoded.specs_for_crossover
        print()
        print("─" * 78)
        print(" SPECS FROM MODEL CODE")
        print("─" * 78)
        key_order = [
            "displacement_cc", "max_pressure_bar", "max_speed_rpm",
            "max_flow_lpm", "max_torque_nm",
            "mount_pref", "shaft_pref", "shaft_diameter_mm",
            "port_pref", "rotation",
        ]
        for k in key_order:
            v = specs.get(k)
            marker = "  " if v not in (None, "") else "○ "
            print(f"  {marker}{k:<22} {_fmt(v)}")

    print()
    print("─" * 78)
    print(" SERIES MAPPING")
    print("─" * 78)
    if result.mapped_hengli_series:
        print(f"  {result.competitor_brand} {result.competitor_series}  →  "
              f"Hengli {result.mapped_hengli_series}")
        if result.mapping_rationale:
            words = result.mapping_rationale.split()
            line = "  "
            for w in words:
                if len(line) + len(w) + 1 > 76:
                    print(line)
                    line = "  " + w
                else:
                    line += (" " if line.strip() else "") + w
            if line.strip():
                print(line)
    else:
        print("  (no mapping resolved)")

    if result.primary:
        print()
        print("═" * 78)
        print(f" PRIMARY HENGLI CANDIDATES ({result.mapped_hengli_series} series)")
        print("═" * 78)
        print()
        for i, m in enumerate(result.primary, 1):
            _print_match_block(m, i, "PRIMARY")

    if result.fallback:
        print()
        print("═" * 78)
        print(f" ALTERNATIVE-SERIES CANDIDATES (other Hengli families)")
        print("═" * 78)
        print()
        for i, m in enumerate(result.fallback, 1):
            _print_match_block(m, i, "ALTERNATIVE")

    all_results = list(result.primary) + list(result.fallback)
    if len(all_results) > 1:
        print("─" * 78)
        print(" COMPARISON SUMMARY")
        print("─" * 78)
        print(f"  {'Tier':<14}{'Part Number':<28}{'Disp':<11}{'Δp cont':<11}"
              f"{'Fit':<14}{'Score'}")
        for m in result.primary:
            print(f"  {'PRIMARY':<14}{m.suggested_part_number:<28}"
                  f"{_fmt(m.model.get('displacement_cc'), ' cc'):<11}"
                  f"{_fmt(m.model.get('max_dp_cont'), ' bar'):<11}"
                  f"{m.fit_rating:<14}{m.score}")
        for m in result.fallback:
            print(f"  {'ALTERNATIVE':<14}{m.suggested_part_number:<28}"
                  f"{_fmt(m.model.get('displacement_cc'), ' cc'):<11}"
                  f"{_fmt(m.model.get('max_dp_cont'), ' bar'):<11}"
                  f"{m.fit_rating:<14}{m.score}")
        print()

    if result.warnings:
        print("─" * 78)
        print(" NOTES & WARNINGS")
        print("─" * 78)
        for w in result.warnings:
            print(f"  ⚠ {w}")
        print()

    print("═" * 78)
    if not result.is_valid:
        print(" ❌ Crossover could not be completed.")
    elif result.primary and result.primary[0].score >= 70:
        print(" ✅ Strong primary match found.")
    elif result.primary:
        print(" ⚠  Primary match available but weaker than ideal.")
    elif result.fallback:
        print(" ⚠  No primary match — alternative-series candidates shown.")
    print("═" * 78)
    print()


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    base_dir = Path(__file__).parent
    for required in ("Hengli_Orbital_Motor_Master.xlsx",
                     "competitor_code_extractor.xlsx",
                     "spec_matching.py",
                     "decode_competitor.py"):
        if not (base_dir / required).exists():
            print(f"❌ Required file missing in script folder: {required}")
            sys.exit(1)

    if len(sys.argv) > 1:
        code = " ".join(sys.argv[1:])
        result = crossover_competitor(code)
        print_crossover(result)
    else:
        print("Competitor → Hengli Crossover")
        print("(enter blank or 'quit' to exit)")
        while True:
            try:
                code = input("\nCompetitor model code: ").strip()
            except (EOFError, KeyboardInterrupt):
                print()
                break
            if not code or code.lower() in ("quit", "exit", "q"):
                break
            result = crossover_competitor(code)
            print_crossover(result)
