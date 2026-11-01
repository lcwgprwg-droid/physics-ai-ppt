from pathlib import Path
import json, sys

ROOT = Path(__file__).resolve().parents[1] / "examples" / "M03-microelement-velocity"
spec = json.loads((ROOT / "course_spec.json").read_text(encoding="utf-8"))
coverage = json.loads((ROOT / "note_coverage_matrix.json").read_text(encoding="utf-8"))
baseline = json.loads((ROOT / "golden_baseline.json").read_text(encoding="utf-8"))

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
    "diagram_first_layout",
    "native_shape_for_fragile_diagrams",
    "no_missing_media_relationships",
    "no_out_of_bounds_objects",
    "ppt_render_compatibility",
    "full_deck_contact_sheet_review",
}
missing_gates = required_gates - set(spec.get("golden_gates", []))
if missing_gates:
    errors.append(f"missing golden gates: {sorted(missing_gates)}")

for item in coverage:
    if item.get("stem_status") == "diagram_and_derivation_only" and item.get("confidence") != "medium":
        errors.append(f"{item.get('id')}: source-gap example must remain explicitly marked")

stable = spec.get("stable_baseline", {})
if stable.get("version") != "V1.1":
    errors.append("M03 stable baseline must be V1.1")
if stable.get("slide_count") != 38:
    errors.append("M03 V1.1 stable baseline must remain 38 slides")
if stable.get("missing_media_relationships") != 0:
    errors.append("stable baseline cannot contain missing media relationships")
if stable.get("out_of_bounds_objects") != 0:
    errors.append("stable baseline cannot contain out-of-bounds objects")
if sorted(stable.get("native_shape_risk_pages", [])) != [1, 7, 9]:
    errors.append("native-shape compatibility pages must remain [1, 7, 9]")

expected = baseline.get("expected", {})
if expected.get("slide_count") != 38:
    errors.append("golden baseline expected slide_count must be 38")
if expected.get("required_examples") != 6:
    errors.append("golden baseline must require all 6 examples")
if expected.get("required_note_items") != 14:
    errors.append("golden baseline must require all 14 note coverage items")
if expected.get("missing_media_relationships") != 0:
    errors.append("golden baseline expects zero missing media relationships")
if expected.get("out_of_bounds_objects") != 0:
    errors.append("golden baseline expects zero out-of-bounds objects")
if expected.get("render_qa") != "passed":
    errors.append("golden baseline render QA must be passed")
if expected.get("contact_sheet_review") != "passed":
    errors.append("golden baseline contact-sheet review must be passed")

fingerprints = baseline.get("artifact_fingerprints", {})
for artifact in ("pptx", "rendered_pdf", "contact_sheet"):
    if not fingerprints.get(artifact, {}).get("sha256"):
        errors.append(f"missing SHA-256 fingerprint for {artifact}")

if errors:
    print("FAIL")
    for e in errors:
        print("-", e)
    sys.exit(1)

print("PASS")
print(f"{len(actual_ids)} required note items covered")
print("6 examples preserved; source-gap examples explicitly marked")
print("M03 V1.1 locked as 38-slide stable Golden Baseline")
print("compatibility gates: media=0 missing, bounds=0, render=passed, native pages=[1,7,9]")
