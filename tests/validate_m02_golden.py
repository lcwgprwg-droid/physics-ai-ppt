from pathlib import Path
import json, re, sys

ROOT = Path(__file__).resolve().parents[1] / "examples" / "M02-calculus"
spec = json.loads((ROOT/"course_spec.json").read_text(encoding="utf-8"))
forms = json.loads((ROOT/"formula_registry.json").read_text(encoding="utf-8"))

errors=[]

if spec["duration_min"] != 90:
    errors.append("duration_min must be 90 for this golden test")

roles={f["role"] for f in forms}
required={"instantaneous_velocity","velocity_derivative","acceleration_derivative",
          "finite_increment","differential_relation","displacement_integral","distance"}
missing=required-roles
if missing:
    errors.append(f"missing formula roles: {sorted(missing)}")

for f in forms:
    latex=f["latex"]
    if re.search(r"(?<!\\mathrm\{)d[xtv]", latex):
        errors.append(f"{f['id']}: possible non-upright differential operator")
    if f["role"]=="displacement_integral" and "\\tau" in latex:
        errors.append("M02 classroom displacement integral should use t, not tau")

if errors:
    print("FAIL")
    for e in errors:
        print("-",e)
    sys.exit(1)

print("PASS")
print(f"{len(forms)} formulas checked")
print("M02 notation decision: t is used consistently for classroom integral notation")
