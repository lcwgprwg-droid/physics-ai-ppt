# Physics AI Courseware Builder

AI-assisted production pipeline for high-school physics competition / Strong Foundation courseware.

The project turns teacher handwritten notes or accepted lesson masters into a reproducible course package while preserving the teacher’s own lesson logic.

## Core principle

**Teacher notes / accepted lesson master define the lesson boundary.**

If a source set is intended as one 90-minute lesson, the system must cover its meaningful content at 100% before release. External university physics / Olympiad / GitHub sources may verify or improve presentation, but they may not silently delete or replace source content.

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
- prefer native PowerPoint shapes for fragile/core diagrams
- student-facing classroom content stays in the main flow
- teacher source/QA metadata goes to appendix when appropriate
- full-deck PDF/PNG render before release

## Physics Courseware Skill

Skill entry:

`skills/physics-courseware-builder/SKILL.md`

Supporting files:

- `skills/physics-courseware-builder/references/WORKFLOW.md`
- `skills/physics-courseware-builder/references/QUALITY_GATES.md`
- `skills/physics-courseware-builder/prompts/MASTER_PROMPT.md`

## Golden Baselines

### M01 — 向量

Accepted baseline: **V2.0 稳定重制版**

Focus:
- scalar/vector distinction
- position vector and displacement
- vector-translation boundary
- components/unit vectors/2D/3D
- addition/subtraction/resultant range
- decomposition/projection
- dot product/work/cosine-law derivation
- cross product/area/torque entry
- student-facing main flow + teacher appendix
- native PowerPoint geometry for compatibility

Files:
- `examples/M01-vector/course_spec.json`
- `examples/M01-vector/coverage_matrix.json`
- `examples/M01-vector/golden_baseline.json`
- `tests/validate_m01_golden.py`

Run:

```bash
python tests/validate_m01_golden.py
```

### M02 — 数学工具与连续运动的语言

The original handwritten note set is treated as one 90-minute lesson and covers vector/trig tools, coordinate systems, limits, derivatives, integrals, summation and the ant-climbing-tree example.

Files:
- `examples/M02-calculus/course_spec.json`
- `examples/M02-calculus/formula_registry.json`
- `examples/M02-calculus/note_coverage_matrix.json`
- `tests/validate_m02_golden.py`

Run:

```bash
python tests/validate_m02_golden.py
```

### M03 — 微元法求速度

Accepted baseline: **V1.1 全页稳定版**

Focus:
- 3-page source set preserved as one 90-minute lesson
- 6 examples preserved
- geometric redraws
- micro-displacement method
- small-angle relation
- sine-law `2R`, cosine law and circumcircle bridge
- native-shape compatibility pages for fragile diagrams

Files:
- `examples/M03-microelement-velocity/course_spec.json`
- `examples/M03-microelement-velocity/note_coverage_matrix.json`
- `examples/M03-microelement-velocity/golden_baseline.json`
- `tests/validate_m03_golden.py`

Run:

```bash
python tests/validate_m03_golden.py
```

## Regression contract

A renderer / skill / layout change is not considered safe unless it preserves all three accepted baselines:

```bash
python tests/validate_m01_golden.py
python tests/validate_m02_golden.py
python tests/validate_m03_golden.py
```

The three baselines intentionally stress different failure modes:

- **M01:** vector geometry, native diagrams, student-facing lesson structure
- **M02:** formula typography, calculus, complete stems, coordinate systems
- **M03:** complex handwritten geometry, microelement derivations, image compatibility

## Roadmap

1. Freeze specification and QA contract — **done**
2. M01 / M02 / M03 Golden Baselines — **done**
3. Build `course_spec -> renderer -> PPT/PDF -> QA` automation
4. Add reusable semantic layout components
5. Add automated overlap/media/render checks to CI
6. Use M04 as the next unseen-course production test

## Repository status

Current Golden-Baseline work is being reviewed on dedicated test branches before integration into the default branch.
