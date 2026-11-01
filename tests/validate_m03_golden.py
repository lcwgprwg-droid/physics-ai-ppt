from pathlib import Path
import json, sys

ROOT = Path(__file__).resolve().parents[1] / "examples" / "M03-microelement-velocity"
spec = json.loads((ROOT / "course_spec.json").read_text(encoding="utf-8"))
coverage = json.loads((ROOT / "note_coverage_matrix.json").read_text(encoding="utf-8"))

errors = []

if spec.get("duration_min") != 90:
    errors.append("M03 must remain a single 90-minute lesson")
if "不得擅自拆课" not in spec.get("source_boundary", ""):
    errors.append("source boundary rule is missing")
if spec.get("problem_policy", {}).get("examples_count") != 6:
    errors.append("M03 must preserve all 6 note examples")

required_ids = {
    "P1-01","P1-02","P1-03","P1-04","P1-05","P1-06",
    "P2-01","P2-02","P2-03","P2-04",
    "P3-01","P3-02","P3-03","P3-04"
}
actual_ids = {item.get("id") for item in coverage if item.get("status") == "required"}
missing = required_ids - actual_ids
if missing:
    errors.append(f"missing required note items: {sorted(missing)}")

required_gates = {
    "notes_coverage_100_percent",
    "all_6_examples_present",
    "all_source_diagrams_redrawn",
    "small_angle_relation_present",
    "sine_law_2R_present",
    "cosine_law_present",
    "example6_circumcircle_bridge_present",
    "formula_safe_zone",
    "ppt_render_compatibility",
}
missing_gates = required_gates - set(spec.get("golden_gates", []))
if missing_gates:
    errors.append(f"missing golden gates: {sorted(missing_gates)}")

for item in coverage:
    if item.get("stem_status") == "diagram_and_derivation_only" and item.get("confidence") != "medium":
        errors.append(f"{item.get('id')}: source-gap example must remain explicitly marked")

if errors:
    print("FAIL")
    for e in errors:
        print("-", e)
    sys.exit(1)

print("PASS")
print(f"{len(actual_ids)} required note items covered")
print("6 examples preserved; source-gap examples explicitly marked")
print("M03 is ready for renderer/slide-generation pressure testing")
