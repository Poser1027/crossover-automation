"""
Hengli Orbital Motor Part Number Decoder
=========================================

Reverse-decode a Hengli orbital motor part number into its full
specifications, mount/port/shaft descriptions, and option meanings.

Validates the part number against the master workbook and flags:
  - Unknown codes (typos, codes not in catalog)
  - Invalid combinations (e.g. bearingless mount with non-Cardan shaft)
  - Missing/extra positions

USAGE
-----
    # As CLI:
    python decode_part_number.py HSD-080-A3-H-S1-A-N-A
    python decode_part_number.py "HSD 080 A3 H S1 A N A"
    python decode_part_number.py HSD080A3HS1ANA
    python decode_part_number.py            # interactive prompt

    # Programmatic:
    from decode_part_number import decode_part_number, print_decoded
    result = decode_part_number("HSD-080-A3-H-S1-A-N-A", db)
    print_decoded(result)

REQUIREMENTS
------------
- Python 3.9+
- openpyxl
- spec_matching.py in the same folder (re-uses MotorDatabase)
- Hengli_Orbital_Motor_Master.xlsx alongside the scripts
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import re
import sys

# Re-use the database loader from the spec-matching tool
from spec_matching import MotorDatabase


# ─────────────────────────────────────────────────────────────────────────────
# Part-number structure per series
# ─────────────────────────────────────────────────────────────────────────────
SERIES_LAYOUT = {
    # HRD: 7 positions — series, type, mount/port (combined), shaft, rotation, paint, special
    "HRD": ["series", "type", "mount_port", "shaft", "rotation", "paint", "special"],
    # HSD/HSP/HSE: 8 positions — series, type, mount, port, shaft, rotation, paint, special
    "HSD": ["series", "type", "mount", "port", "shaft", "rotation", "paint", "special"],
    "HSP": ["series", "type", "mount", "port", "shaft", "rotation", "paint", "special"],
    "HSE": ["series", "type", "mount", "port", "shaft", "rotation", "paint", "special"],
}

KNOWN_SERIES = list(SERIES_LAYOUT.keys())

# Field widths for tokenising a concatenated part number (e.g. HSD080A3HS1ANA).
# Type is always 3 digits. Mount/port/shaft codes vary in length so we need a
# table of valid codes per position to split correctly.

TYPE_REGEX = re.compile(r"^\d{3}$")


# ─────────────────────────────────────────────────────────────────────────────
# Result containers
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class DecodedField:
    """A single decoded position from the part number."""
    position: str           # e.g. "series", "mount", "shaft"
    raw: str                # the raw code as input by the user
    recognised: bool        # could we look it up?
    description: str = ""   # human-readable description
    extra: dict = field(default_factory=dict)  # additional fields (e.g. port thread)


@dataclass
class DecodedPartNumber:
    """Full decoded result."""
    raw_input: str
    normalised: str         # cleaned, dash-separated form
    series: Optional[str]   # 'HRD', 'HSD', etc.
    fields: dict            # position name -> DecodedField
    model: Optional[dict]   # the matching row from the Models sheet, if found
    unknown_fields: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    is_valid: bool = True   # False if the structure couldn't be parsed at all


# ─────────────────────────────────────────────────────────────────────────────
# Tokenisation
# ─────────────────────────────────────────────────────────────────────────────
def _split_part_number(raw: str, db: MotorDatabase) -> tuple[Optional[str], list[str], str]:
    """
    Break a raw part number into tokens. Handles dash, space, or no separator.

    Returns (series, tokens, error_message). Series is None if not recognised.
    """
    if not raw or not raw.strip():
        return None, [], "Empty input."

    s = raw.strip().upper()

    # First, try simple separator-based splits (dash, space, slash, dot)
    tokens = re.split(r"[-\s/.]+", s)
    tokens = [t for t in tokens if t]

    if tokens and tokens[0] in KNOWN_SERIES:
        # If splitting by separators gave us a known series prefix, trust it.
        return tokens[0], tokens, ""

    # Otherwise the input is probably concatenated (e.g. HSD080A3HS1ANA).
    # Find the series prefix.
    series = None
    for s_candidate in KNOWN_SERIES:
        if s.startswith(s_candidate):
            series = s_candidate
            break

    if series is None:
        return None, [], (
            f"Could not identify series. Expected one of {KNOWN_SERIES}, "
            f"got: {s[:6]!r}"
        )

    # Greedy-tokenise the rest using known code tables for this series.
    rest = s[len(series):]
    tokens = _tokenise_concatenated(series, rest, db)
    tokens.insert(0, series)
    return series, tokens, ""


def _tokenise_concatenated(series: str, rest: str, db: MotorDatabase) -> list[str]:
    """
    Split a concatenated tail like '080A3HS1ANA' into ordered tokens
    using the known codes for this series.
    """
    tokens = []

    # Type is always 3 digits
    if len(rest) >= 3 and rest[:3].isdigit():
        tokens.append(rest[:3])
        rest = rest[3:]
    else:
        return tokens  # malformed

    # Build sets of valid codes for each position from the database
    mount_codes = {m["code"] for m in db.mounts_for_series(series) if m["code"]}
    port_codes = {p["code"] for p in db.ports_for_series(series) if p["code"]}
    shaft_codes = {s["code"] for s in db.shafts_for_series(series) if s["code"]}

    rotation_codes = {"A", "R"}
    paint_codes = {"N", "B", "C"}
    # Special features: anything in the catalog plus the codes the user might type.
    special_codes = {
        o["code"] for o in db.options_for_series(series) if o["position"] == "Special features"
    }

    if series == "HRD":
        # Order: mount_port, shaft, rotation, paint, special
        position_codes_order = [mount_codes, shaft_codes, rotation_codes, paint_codes, special_codes]
    else:
        # Order: mount, port, shaft, rotation, paint, special
        position_codes_order = [mount_codes, port_codes, shaft_codes, rotation_codes, paint_codes, special_codes]

    for valid_codes in position_codes_order:
        if not rest:
            break
        # Match longest valid code first (e.g. 'DC' before 'D')
        sorted_codes = sorted(valid_codes, key=len, reverse=True)
        matched = None
        for code in sorted_codes:
            if rest.startswith(code):
                matched = code
                break
        if matched:
            tokens.append(matched)
            rest = rest[len(matched):]
        else:
            # Couldn't match — take 1-2 chars as a best guess so the user sees what's wrong
            tokens.append(rest[:2] if len(rest) >= 2 else rest)
            rest = rest[2:] if len(rest) >= 2 else ""

    if rest:
        tokens.append(rest)  # trailing junk; will surface as a warning later

    return tokens


# ─────────────────────────────────────────────────────────────────────────────
# Field decoding (one code -> description)
# ─────────────────────────────────────────────────────────────────────────────
def _decode_type(series: str, code: str, db: MotorDatabase) -> DecodedField:
    """Find the model row for series + type code."""
    for m in db.models:
        if m["series"] == series and str(m["type"]) == str(code):
            return DecodedField(
                position="type", raw=code, recognised=True,
                description=f"{series} frame size {code}",
                extra={"model_row": m},
            )
    return DecodedField(position="type", raw=code, recognised=False,
                        description="(unknown frame size)")


def _decode_mount(series: str, code: str, db: MotorDatabase) -> DecodedField:
    for m in db.mounts_for_series(series):
        if m["code"] == code:
            extra = {"port": m.get("port"), "drain": m.get("drain")} if m.get("port") else {}
            return DecodedField(position="mount", raw=code, recognised=True,
                                description=m["description"], extra=extra)
    return DecodedField(position="mount", raw=code, recognised=False,
                        description="(unknown mount code)")


def _decode_port(series: str, code: str, db: MotorDatabase) -> DecodedField:
    for p in db.ports_for_series(series):
        if p["code"] == code:
            return DecodedField(
                position="port", raw=code, recognised=True,
                description=f"Port A/B: {p['port_ab']}, drain: {p['drain']}"
                            + (f"  ({p['notes']})" if p.get("notes") else ""),
                extra={"port_ab": p["port_ab"], "drain": p["drain"], "notes": p.get("notes", "")},
            )
    return DecodedField(position="port", raw=code, recognised=False,
                        description="(unknown port code)")


def _decode_shaft(series: str, code: str, db: MotorDatabase) -> DecodedField:
    for s in db.shafts_for_series(series):
        if s["code"] == code:
            return DecodedField(
                position="shaft", raw=code, recognised=True,
                description=s["description"],
                extra={
                    "max_torque_nm": s.get("max_torque_nm"),
                    "compatible_mounts": s.get("compatible_mounts", ""),
                },
            )
    return DecodedField(position="shaft", raw=code, recognised=False,
                        description="(unknown shaft code)")


def _decode_rotation(code: str) -> DecodedField:
    mapping = {"A": "CW (clockwise)", "R": "CCW (counter-clockwise)"}
    desc = mapping.get(code)
    return DecodedField(position="rotation", raw=code,
                        recognised=desc is not None,
                        description=desc or "(unknown rotation code)")


def _decode_paint(code: str) -> DecodedField:
    mapping = {"N": "No paint", "B": "Black", "C": "Hengli blue"}
    desc = mapping.get(code)
    return DecodedField(position="paint", raw=code,
                        recognised=desc is not None,
                        description=desc or "(unknown paint code)")


def _decode_special(series: str, code: str, db: MotorDatabase) -> DecodedField:
    """Look up Special features for this series."""
    for o in db.options_for_series(series):
        if o["position"] == "Special features" and o["code"] == code:
            desc = o["description"]
            if o.get("notes"):
                desc += f"  ({o['notes']})"
            return DecodedField(position="special", raw=code, recognised=True, description=desc)
    return DecodedField(position="special", raw=code, recognised=False,
                        description="(unknown special-feature code)")


# ─────────────────────────────────────────────────────────────────────────────
# Combination validation
# ─────────────────────────────────────────────────────────────────────────────
def _validate_combinations(series: str, fields: dict, model: Optional[dict]) -> list[str]:
    """Check for invalid code combinations. Returns a list of warning strings."""
    warnings = []

    mount = fields.get("mount") or fields.get("mount_port")
    shaft = fields.get("shaft")
    special = fields.get("special")

    # --- Mount/shaft compatibility ---
    if shaft and shaft.recognised and shaft.extra.get("compatible_mounts"):
        compat = shaft.extra["compatible_mounts"]
        mount_code = mount.raw if mount else None
        if mount_code:
            compat_l = compat.lower()
            mount_l = mount_code.lower()
            compatible = False

            if "all" in compat_l and "except" not in compat_l:
                compatible = True
            elif "except" in compat_l:
                # "All HSP mounts EXCEPT W1/W4" — compatible if mount NOT in the exclusion list
                exclusion = compat_l.split("except", 1)[1]
                compatible = mount_l not in exclusion
            elif mount_code in compat or mount_l in compat_l:
                compatible = True
            elif "wheel" in compat_l and mount_code in ("W1", "W4"):
                compatible = True

            if not compatible:
                warnings.append(
                    f"Shaft {shaft.raw} is not compatible with mount {mount_code}. "
                    f"Compatible mounts: {compat}"
                )

    # --- Bearingless mounts (HSD B0/F0) require Cardan shaft C1 ---
    if series == "HSD" and mount and mount.raw in ("B0", "F0"):
        if shaft and shaft.raw != "C1":
            warnings.append(
                f"Mount {mount.raw} is a bearingless mount and requires Cardan shaft C1; "
                f"got shaft {shaft.raw}."
            )

    # --- HSD C1 shaft (Cardan) requires bearingless mount ---
    if series == "HSD" and shaft and shaft.raw == "C1":
        if mount and mount.raw not in ("B0", "F0"):
            warnings.append(
                f"Cardan shaft C1 requires a bearingless mount (B0 or F0); "
                f"got mount {mount.raw}."
            )

    # --- HSP wheel-only shafts (TN, T4) with non-wheel mount ---
    if series == "HSP" and shaft and shaft.raw in ("TN", "T4"):
        if mount and mount.raw not in ("W1", "W4"):
            warnings.append(
                f"Shaft {shaft.raw} is wheel-mount only (W1/W4); got mount {mount.raw}."
            )

    # --- HSP non-wheel shafts paired with wheel mount ---
    if series == "HSP" and mount and mount.raw in ("W1", "W4"):
        if shaft and shaft.raw not in ("TN", "T4"):
            warnings.append(
                f"Wheel mount {mount.raw} requires a Wheel shaft (TN or T4); "
                f"got shaft {shaft.raw}."
            )

    # --- HRD speed-sensor shafts (SP, SQ, SV, ST, SY) require sensor-prep mount (A29/A30/A66) ---
    if series == "HRD" and shaft and shaft.raw in ("SP", "SQ", "SV", "ST", "SY"):
        if mount and mount.raw not in ("A29", "A30", "A66"):
            warnings.append(
                f"Shaft {shaft.raw} is for sensor-prep mounts (A29/A30/A66) only; "
                f"got mount {mount.raw}."
            )

    # --- HRD speed-sensor special feature (S2) requires sensor-prep mount ---
    if series == "HRD" and special and special.raw == "S2":
        if mount and mount.raw not in ("A29", "A30", "A66"):
            warnings.append(
                f"Special feature S2 (speed sensor) requires mounts A29/A30/A66; "
                f"got mount {mount.raw}."
            )

    # --- HSP dust cover (D) on Magneto/Square/Wheel mounts requires special handling ---
    # This isn't strictly invalid — catalog says "available on request" for F6/M0/W1.
    # We don't flag this as a warning, just informational at most.

    return warnings


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────
def decode_part_number(raw: str, db: MotorDatabase) -> DecodedPartNumber:
    """
    Decode a Hengli orbital motor part number into its full specs and field meanings.

    Returns a DecodedPartNumber object. Even on failure, the object describes what
    could and couldn't be parsed (via is_valid, unknown_fields, warnings).
    """
    result = DecodedPartNumber(
        raw_input=raw, normalised="", series=None, fields={}, model=None,
    )

    series, tokens, err = _split_part_number(raw, db)
    if err:
        result.is_valid = False
        result.warnings.append(err)
        return result

    result.series = series
    layout = SERIES_LAYOUT[series]
    expected_positions = len(layout)

    # Check token count
    if len(tokens) < expected_positions:
        result.warnings.append(
            f"Part number has {len(tokens)} positions; "
            f"{series} expects {expected_positions}. "
            f"Missing trailing positions will be marked unrecognised."
        )
    elif len(tokens) > expected_positions:
        result.warnings.append(
            f"Part number has {len(tokens)} positions; "
            f"{series} expects {expected_positions}. "
            f"Extra tokens: {tokens[expected_positions:]}"
        )

    # Build normalised string
    result.normalised = "-".join(tokens[:expected_positions]
                                 + tokens[expected_positions:])

    # Decode each position
    for i, position_name in enumerate(layout):
        if i >= len(tokens):
            # Missing position
            field_obj = DecodedField(position=position_name, raw="",
                                     recognised=False,
                                     description="(missing from part number)")
            result.fields[position_name] = field_obj
            result.unknown_fields.append(position_name)
            continue

        code = tokens[i]

        if position_name == "series":
            field_obj = DecodedField(position="series", raw=code,
                                     recognised=(code in KNOWN_SERIES),
                                     description=f"{code} series orbital hydraulic motor")
        elif position_name == "type":
            field_obj = _decode_type(series, code, db)
            if field_obj.recognised:
                result.model = field_obj.extra.get("model_row")
        elif position_name in ("mount", "mount_port"):
            field_obj = _decode_mount(series, code, db)
            # Override the position name for HRD so output stays consistent
            field_obj.position = position_name
        elif position_name == "port":
            field_obj = _decode_port(series, code, db)
        elif position_name == "shaft":
            field_obj = _decode_shaft(series, code, db)
        elif position_name == "rotation":
            field_obj = _decode_rotation(code)
        elif position_name == "paint":
            field_obj = _decode_paint(code)
        elif position_name == "special":
            field_obj = _decode_special(series, code, db)
        else:
            field_obj = DecodedField(position=position_name, raw=code,
                                     recognised=False, description="(unhandled position)")

        result.fields[position_name] = field_obj
        if not field_obj.recognised:
            result.unknown_fields.append(position_name)

    # Run combination validation
    result.warnings.extend(_validate_combinations(series, result.fields, result.model))

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


def print_decoded(decoded: DecodedPartNumber) -> None:
    """Pretty-print a decoded part number with full specs and warnings."""
    print()
    print("═" * 78)
    print(" HENGLI PART NUMBER DECODER")
    print("═" * 78)
    print(f"\n  Input:      {decoded.raw_input}")
    print(f"  Normalised: {decoded.normalised or '(could not parse)'}")

    if not decoded.is_valid:
        print(f"\n  ❌ Could not decode this part number.")
        for w in decoded.warnings:
            print(f"     {w}")
        print()
        return

    print(f"  Series:     {decoded.series}")
    print()

    # ─── Decoded fields ───
    print("─" * 78)
    print(" DECODED FIELDS")
    print("─" * 78)
    label_map = {
        "series": "Series",
        "type": "Type / Frame",
        "mount": "Mount",
        "mount_port": "Mount + Port (HRD combined)",
        "port": "Port",
        "shaft": "Shaft",
        "rotation": "Rotation",
        "paint": "Paint",
        "special": "Special features",
    }
    for pos_name, field_obj in decoded.fields.items():
        marker = "  " if field_obj.recognised else "⚠ "
        label = label_map.get(pos_name, pos_name)
        code_str = field_obj.raw or "(missing)"
        print(f"  {marker}{label:<32} {code_str:<6}  {field_obj.description}")
        if field_obj.extra:
            for k, v in field_obj.extra.items():
                if k == "model_row":
                    continue
                if v not in (None, ""):
                    print(f"           {k:<28}     {v}")

    # ─── Model specs ───
    if decoded.model:
        m = decoded.model
        print()
        print("─" * 78)
        print(" MODEL SPECIFICATIONS")
        print("─" * 78)
        print(f"  Brand / Series:       {m.get('brand')} / {m.get('series')}")
        print(f"  Frame size:           {m.get('type')}")
        print(f"  Distribution:         {m.get('distribution') or '—'}")
        print(f"  Displacement:         {_fmt(m.get('displacement_cc'))} cc/rev")
        print()
        print(f"  RATINGS                     Continuous   Intermittent   Peak")
        print(f"  Max speed (rpm)             {_fmt(m.get('max_speed_cont')):>10}   "
              f"{_fmt(m.get('max_speed_inter')):>12}   —")
        print(f"  Max torque (N·m)            {_fmt(m.get('max_torque_cont')):>10}   "
              f"{_fmt(m.get('max_torque_inter')):>12}   —")
        print(f"  Max Δp (bar)                {_fmt(m.get('max_dp_cont')):>10}   "
              f"{_fmt(m.get('max_dp_inter')):>12}   "
              f"{_fmt(m.get('max_dp_peak')):>5}")
        print(f"  Max flow (L/min)            {_fmt(m.get('max_flow_cont')):>10}   "
              f"{_fmt(m.get('max_flow_inter')):>12}   —")
        print(f"  Max output (kW)             {_fmt(m.get('max_output_cont')):>10}   "
              f"{_fmt(m.get('max_output_inter')):>12}   —")
        print()
        print(f"  Min start torque (N·m):  cont @ max Δp = {_fmt(m.get('min_start_torque_cont'))}, "
              f"inter @ max Δp = {_fmt(m.get('min_start_torque_inter'))}")
        print(f"  No-load starting pressure: {_fmt(m.get('max_no_load_start_p'))} bar")
        print(f"  Weight:                   standard {_fmt(m.get('weight_standard'))} kg, "
              f"bearingless {_fmt(m.get('weight_bearingless'))} kg")

    # ─── Warnings ───
    if decoded.warnings:
        print()
        print("─" * 78)
        print(" WARNINGS / VALIDATION ISSUES")
        print("─" * 78)
        for w in decoded.warnings:
            print(f"  ⚠ {w}")

    # ─── Unknown fields ───
    if decoded.unknown_fields:
        print()
        print("─" * 78)
        print(" UNRECOGNISED FIELDS")
        print("─" * 78)
        for pf in decoded.unknown_fields:
            f = decoded.fields[pf]
            label = label_map.get(pf, pf)
            print(f"  ⚠ {label} (code '{f.raw}') — {f.description}")

    print()
    print("═" * 78)
    if not decoded.warnings and not decoded.unknown_fields:
        print(" ✅ Valid part number — all fields recognised, no conflicts.")
    elif decoded.unknown_fields:
        print(" ❌ Part number has unrecognised codes.")
    else:
        print(" ⚠  Part number recognised but has validation warnings.")
    print("═" * 78)
    print()


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    xlsx_path = Path(__file__).parent / "Hengli_Orbital_Motor_Master.xlsx"
    if not xlsx_path.exists():
        print(f"❌ Could not find {xlsx_path}")
        print("   Make sure Hengli_Orbital_Motor_Master.xlsx is in the same folder.")
        sys.exit(1)

    db = MotorDatabase(str(xlsx_path))

    if len(sys.argv) > 1:
        # CLI argument provided — decode and exit
        pn = " ".join(sys.argv[1:])
        result = decode_part_number(pn, db)
        print_decoded(result)
    else:
        # Interactive mode
        print("Hengli Part Number Decoder")
        print("(enter blank or 'quit' to exit)")
        while True:
            try:
                pn = input("\nPart number: ").strip()
            except (EOFError, KeyboardInterrupt):
                print()
                break
            if not pn or pn.lower() in ("quit", "exit", "q"):
                break
            result = decode_part_number(pn, db)
            print_decoded(result)
