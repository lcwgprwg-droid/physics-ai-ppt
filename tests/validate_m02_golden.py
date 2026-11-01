from pathlib import Path
import json
import re
import sys

ROOT = Path(__file__).resolve().parents[1] / "examples" / "M02-calculus"
spec = json.loads((ROOT / "course_spec.json").read_text(encoding="utf-8"))
forms = json.loads((ROOT / "formula_registry.json").read_text(encoding="utf-8"))
coverage = json.loads((ROOT / "note_coverage_matrix.json").read_text(encoding="utf-8"))

errors = []

# Lesson contract
if spec.get("golden_test") is not True:
    errors.append("M02 must remain marked as a golden test")
if spec.get("duration_min") != 90:
    errors.append("duration_min must be 90 for this golden test")
if spec.get("lesson_number") != 2:
    errors.append("M02 lesson_number must be 2")
if spec.get("theme") != "数学工具与连续运动的语言":
    errors.append("M02 theme changed unexpectedly")

boundary = spec.get("source_boundary", {})
if boundary.get("required_coverage_percent") != 100:
    errors.append("note coverage target must be 100%")
if boundary.get("allow_silent_split_to_next_lesson") is not False:
    errors.append("silent split-to-next-lesson must be disabled")

# Coverage matrix
items = coverage.get("items", [])
if not items:
    errors.append("note coverage matrix is empty")
else:
    not_covered = [x.get("id") for x in items if x.get("status") != "covered"]
    if not_covered:
        errors.append(f"uncovered note items: {not_covered}")
    pages = {x.get("note_page") for x in items}
    if pages != {1, 2, 3, 4}:
        errors.append(f"expected note pages 1-4, got {sorted(pages)}")

required_note_items = {
    "正弦定理 a/sinA=b/sinB=c/sinC",
    "余弦定理的向量证明：向量差→平方→点乘展开",
    "三角函数加法公式",
    "由 alpha→-alpha 得到减法公式",
    "直角坐标 A(x,y)",
    "极坐标 A(r,theta) 与 x=r cos theta, y=r sin theta",
    "三维坐标系及按对称性选择柱/球坐标的思想",
    "导数的数学/几何/物理三种语言",
    "蚂蚁上树完整题干：L, v0, h, v(h)=Lv0/(L+h), 求总时间",
    "蚂蚁上树离散法：等分→每段时间→代入→求和→取极限",
    "圆内接三角形与扩展正弦定理 =2R",
}
actual_note_items = {x.get("item") for x in items}
missing_note_items = required_note_items - actual_note_items
if missing_note_items:
    errors.append(f"missing golden note items: {sorted(missing_note_items)}")

# Formula registry
roles = {f["role"] for f in forms}
required_roles = {
    "instantaneous_velocity",
    "velocity_derivative",
    "acceleration_derivative",
    "finite_increment",
    "differential_relation",
    "displacement_integral",
    "distance",
}
missing_roles = required_roles - roles
if missing_roles:
    errors.append(f"missing formula roles: {sorted(missing_roles)}")

for f in forms:
    latex = f["latex"]
    if re.search(r"(?<!\\mathrm\{)d[xtv]", latex):
        errors.append(f"{f['id']}: possible non-upright differential operator")
    if f["role"] == "displacement_integral" and "\\tau" in latex:
        errors.append("M02 classroom displacement integral should use t, not tau")

# Stable visual / compatibility rules
hard_rules = set(spec.get("hard_layout_rules", []))
for rule in [
    "禁止超大公式+小字解释",
    "公式安全区不得被批注/正文/图像侵入",
    "多个公式优先纵向分层",
    "可视化概念优先补图而不是放大公式",
    "推导页采用思路→步骤→结论",
    "题干页以完整可读为第一优先级",
]:
    if rule not in hard_rules:
        errors.append(f"missing hard layout rule: {rule}")

compat = spec.get("compatibility_gates", {})
if compat.get("missing_media_targets") != 0:
    errors.append("golden compatibility target requires zero missing media")
if compat.get("out_of_bounds_objects") != 0:
    errors.append("golden compatibility target requires zero out-of-bounds objects")
if compat.get("full_deck_render_required") is not True:
    errors.append("full deck render must be required")

if spec.get("release_status") != "GOLDEN":
    errors.append("M02 stable spec must remain GOLDEN")

if errors:
    print("FAIL")
    for e in errors:
        print("-", e)
    sys.exit(1)

print("PASS")
print(f"{len(forms)} formulas checked")
print(f"{len(items)} note items checked: 100% covered")
print("M02 golden invariants: lesson boundary, visuals, derivations, problem stems and compatibility gates are locked")
