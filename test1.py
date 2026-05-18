import re
from spec_matching import MotorDatabase, _normalise

db = MotorDatabase("Hengli_Orbital_Motor_Master.xlsx")
candidates = db.shafts_for_series("HSP")
print("初始:", [s['code'] for s in candidates])

# 1. mount 兼容性过滤 (mount_code="A3")
mount_code = "A3"
compatible = []
for s in candidates:
    cm = s["compatible_mounts"] or ""
    cm_l = cm.lower()
    if "all" in cm_l and "except" not in cm_l:
        compatible.append(s)
    elif "except" in cm_l:
        if mount_code.lower() not in cm_l.split("except")[1]:
            compatible.append(s)
    elif mount_code in cm:
        compatible.append(s)
    elif mount_code.lower() in cm_l:
        compatible.append(s)
candidates = compatible
print("after mount:", [s['code'] for s in candidates])

# 2. shaft_pref 过滤 (user_shaft_pref="parallel key")
filtered = [s for s in candidates if "parallel key" in _normalise(s["description"]) or "key" in _normalise(s["description"])]
print("after type:", [s['code'] for s in filtered])
candidates = filtered

# 3. 直径过滤
user_shaft_dia = 31.75
exact_match = []
loose_match = []
for s in candidates:
    m = re.search(r"(?:ø\s*)?([0-9]+(?:\.[0-9]+)?)\s*(?:mm|\(|straight|spline|taper)", _normalise(s["description"]))
    if m:
        dia = float(m.group(1))
        print(f"  {s['code']}: extracted dia = {dia}")
        if abs(dia - user_shaft_dia) <= 0.1:
            exact_match.append(s)
        elif abs(dia - user_shaft_dia) <= 1.0:
            loose_match.append(s)
print("exact:", [s['code'] for s in exact_match])
print("loose:", [s['code'] for s in loose_match])