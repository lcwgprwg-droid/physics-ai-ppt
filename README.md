# Physics AI Courseware Builder

AI-assisted production pipeline for high-school physics competition / Strong Foundation courseware.

The project turns teacher handwritten notes (Notability/PDF/images) into a reproducible course package while preserving the teacher’s own lesson logic.

## Core principle

**Teacher notes are the lesson skeleton, not disposable input.**

If a note set is intended as one 90-minute lesson, the system must cover the meaningful content in that note set at 100% before release. External university physics / Olympiad / GitHub sources may verify or improve presentation, but they may not silently delete or replace teacher-note content.

## What the pipeline produces

- classroom PPT / PDF
- teacher manuscript
- student handout
- Xournal++ board template
- HTML / optional Manim visualization
- layered exercises and problem bank
- formula registry
- note coverage matrix
- QA report

## Production pipeline

```text
Teacher notes / Notability / PDF
        ↓
page-by-page source reconstruction
        ↓
note_coverage_matrix.json   (must reach 100%)
        ↓
course_spec.json + formula_registry.json
        ↓
physics/math QA
        ↓
semantic slide planning
        ↓
visual-first + formula-aware rendering
        ↓
PPT / handout / board / exercises / visualization
        ↓
media + layout + render compatibility QA
        ↓
RELEASE
```

## Stable design rules

- 16:9 classroom projection layout
- one dominant teaching purpose per slide
- no “giant formula + tiny text”
- formula safe zones: no overlap with text/images/annotations
- visual-first explanation for geometric/graphical concepts
- derivations use `idea -> steps -> conclusion`
- one main formula per derivation step when possible
- complete problem stems before solution slides
- formal equations use standard math typography
- handwritten style is reserved for annotations/arrows/highlights
- self-contained media; full-deck PDF/PNG render before release

## Physics Courseware Skill

Skill entry:

`skills/physics-courseware-builder/SKILL.md`

Supporting files:

- `skills/physics-courseware-builder/references/WORKFLOW.md`
- `skills/physics-courseware-builder/references/QUALITY_GATES.md`
- `skills/physics-courseware-builder/prompts/MASTER_PROMPT.md`

## M02 Golden Test

The first accepted baseline is **M02 — 高中物理竞赛第二节：数学工具与连续运动的语言**.

The original four-page handwritten note set is treated as one 90-minute lesson and includes:

- dot/cross product and vector tools
- radians
- sine/cosine laws and vector proof
- trig addition/subtraction formulas
- Cartesian / polar / 3D coordinate-system discussion
- position, displacement, average/instantaneous velocity
- limit, derivative, `x -> v -> a`
- differential/integral language
- Riemann/area interpretation
- summation formulas
- ant-climbing-tree microelement example using both integral and discrete-limit routes
- circumcircle form of the sine law

Golden configuration:

- `examples/M02-calculus/course_spec.json`
- `examples/M02-calculus/formula_registry.json`
- `examples/M02-calculus/note_coverage_matrix.json`
- `tests/validate_m02_golden.py`

Run the Golden validator after modifying the skill/specification:

```bash
python tests/validate_m02_golden.py
```

A stable release must not regress M02 note coverage, mathematical notation, layout rules, problem-stem completeness, or compatibility requirements.

## Roadmap

1. Freeze specification and QA contract — **done**
2. M02 Golden Test — **done**
3. Second pressure test with a substantially different course (M03 microelement method or M01 vectors)
4. Build `course_spec -> renderer -> PPT/PDF -> QA` automation
5. Add reusable semantic layout components and automated overlap/media checks

## Repository status

The current skill/golden-test work lives on branch:

`test/physics-courseware-builder-m02`

See the active pull request for review and integration status.
