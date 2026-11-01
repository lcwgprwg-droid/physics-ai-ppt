from pathlib import Path
import json, sys

ROOT = Path(__file__).resolve().parents[1] / "examples" / "M01-vector"
spec = json.loads((ROOT / "course_spec.json").read_text(encoding="utf-8"))
coverage = json.loads((ROOT / "coverage_matrix.json").read_text(encoding="utf-8"))
baseline = json.loads((ROOT / "golden_baseline.json").read_text(encoding="utf-8"))

errors = []

if spec.get("duration_min") != 90:
    errors.append("M01 must remain a single 90-minute lesson")

stable = spec.get("stable_baseline", {})
if stable.get("version") != "V2.0":
    errors.append("M01 stable baseline must be V2.0")
if stable.get("slide_count") != 40:
    errors.append("M01 V2.0 must remain a 40-slide baseline")
if stable.get("classroom_slides") != "1-37":
    errors.append("student-facing classroom flow must remain slides 1-37")
if stable.get("teacher_appendix") != "38-40":
    errors.append("teacher appendix must remain slides 38-40")
if stable.get("missing_media_relationships") != 0:
    errors.append("M01 stable baseline cannot contain missing media relationships")
if stable.get("out_of_bounds_objects") != 0:
    errors.append("M01 stable baseline cannot contain out-of-bounds objects")
if stable.get("render_qa") != "passed":
    errors.append("M01 stable baseline render QA must be passed")

required_content = set(spec.get("required_content", []))
if len(required_content) != 17:
    errors.append("M01 must preserve all 17 required content items")

required_ids = {f"M01-{i:02d}" for i in range(1, 11)}
actual_ids = {item.get("id") for item in coverage if item.get("status") == "required"}
missing_ids = required_ids - actual_ids
if missing_ids:
    errors.append(f"missing required coverage rows: {sorted(missing_ids)}")

if spec.get("problem_policy", {}).get("classroom_short_exercises") != 3:
    errors.append("M01 must preserve 3 complete classroom short exercises")

required_gates = {
    "content_coverage_100_percent",
    "single_90_minute_lesson",
    "three_complete_classroom_exercises",
    "position_vector_boundary_correct",
    "force_torque_reference_point_dependency_present",
    "dot_and_cross_physical_meaning_present",
    "formula_safe_zone",
    "native_shape_for_core_diagrams",
    "no_missing_media_relationships",
    "no_out_of_bounds_objects",
    "ppt_render_compatibility",
    "full_deck_contact_sheet_review",
}
missing_gates = required_gates - set(spec.get("golden_gates", []))
if missing_gates:
    errors.append(f"missing golden gates: {sorted(missing_gates)}")

expected = baseline.get("expected", {})
if expected.get("slide_count") != 40:
    errors.append("golden baseline expected slide_count must be 40")
if expected.get("lesson_duration_min") != 90:
    errors.append("golden baseline must remain 90 minutes")
if expected.get("classroom_short_exercises") != 3:
    errors.append("golden baseline must require all 3 classroom exercises")
if expected.get("required_content_items") != 17:
    errors.append("golden baseline must require all 17 content items")
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
print("M01 V2.0 locked as 40-slide stable Golden Baseline")
print("classroom slides=1-37; teacher appendix=38-40")
print("17 required content items; 3 complete classroom exercises")
print("compatibility gates: media=0 missing, bounds=0, render=passed")
