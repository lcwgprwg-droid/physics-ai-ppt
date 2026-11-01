# Physics AI Courseware Builder

AI-assisted production pipeline for high-school physics competition / Strong Foundation courseware.

The project turns teacher handwritten notes or accepted lesson masters into a reproducible course package while preserving the teacher’s own lesson logic.

## Quick start

Upload one lesson of handwritten notes / Notability PDF, then say:

> 按竞赛课件 Skill 生成下一节。

Recommended full invocation:

> 按 Physics Courseware Builder 标准生成下一节90分钟竞赛课，严格按我的笔记，不遗漏，不擅自补题干，按 M01–M04 稳定版标准完成 PPT、PDF、覆盖矩阵和全页 QA。

Detailed usage:

`skills/physics-courseware-builder/USAGE.md`

## Core principle

**Teacher notes / accepted lesson master define the lesson boundary.**

If a source set is intended as one 90-minute lesson, the system must cover its meaningful content at 100% before release. External university physics / Olympiad / GitHub sources may verify or improve presentation, but they may not silently delete or replace source content.

If a handwritten example does not contain a complete text stem, the system must preserve it as a diagram-derived example and mark the source gap rather than fabricate missing conditions.

## Stable design rules

- 16:9 classroom projection layout
- one dominant teaching purpose per slide
- no “giant formula + tiny text”
- formula safe zones: no overlap with text/images/annotations
- visual-first explanation for geometric/graphical concepts
- derivations use `idea -> steps -> conclusion`
- one main formula per derivation step when possible
- complete source problem stems before solution slides
- no fabricated stems when the source itself is incomplete
- formal equations use standard math typography
- handwritten style is reserved for annotations/arrows/highlights
- prefer native PowerPoint shapes for fragile/core diagrams
- student-facing classroom content stays in the main flow
- teacher source/QA metadata goes to appendix when appropriate
- full-deck PDF/PNG render before release
- full-deck contact-sheet review before release

## Physics Courseware Skill

Skill entry:

`skills/physics-courseware-builder/SKILL.md`

Supporting files:

- `skills/physics-courseware-builder/USAGE.md`
- `skills/physics-courseware-builder/references/WORKFLOW.md`
- `skills/physics-courseware-builder/references/QUALITY_GATES.md`
- `skills/physics-courseware-builder/prompts/MASTER_PROMPT.md`

Current skill version: **v1.1.0**

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

Focus:
- formula typography
- coordinate systems
- limits / derivatives / integrals
- complete stems
- note coverage discipline

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

## M04 Production Candidate — 相对运动与关联运动

M04 is the first unseen-course test produced after the M01–M03 baselines stabilized.

It tests:
- carry-over microelement example;
- relative position / relative velocity;
- moving reference frames;
- boat-smoke and runner-wind vector problems;
- source-complete vs source-incomplete problem handling;
- associated-motion rigid constraints;
- full-deck compatibility QA.

Files:
- `examples/M04-relative-associated-motion/course_spec.json`
- `examples/M04-relative-associated-motion/coverage_matrix.json`
- `examples/M04-relative-associated-motion/artifact_fingerprints.json`
- `examples/M04-relative-associated-motion/README.md`

Current result: **40 slides / 90 min / 14 source items / media 0 / bounds 0 / render passed**.

## Regression contract

A renderer / skill / layout change is not considered safe unless it preserves the accepted Golden Baselines:

```bash
python tests/validate_m01_golden.py
python tests/validate_m02_golden.py
python tests/validate_m03_golden.py
```

The baselines intentionally stress different failure modes:

- **M01:** vector geometry, native diagrams, student-facing lesson structure
- **M02:** formula typography, calculus, complete stems, coordinate systems
- **M03:** complex handwritten geometry, microelement derivations, image compatibility
- **M04 candidate:** relative motion notation, wind problems, associated-motion constraints, source-gap discipline

## Roadmap

1. Freeze specification and QA contract — **done**
2. M01 / M02 / M03 Golden Baselines — **done**
3. Skill v1.1 quick invocation + source-gap rules — **done**
4. M04 unseen-course production candidate — **done**
5. Promote M04 after teacher acceptance
6. Build `course_spec -> renderer -> PPT/PDF -> QA` automation
7. Add reusable semantic layout components
8. Add automated overlap/media/render checks to CI

## Repository status

Current courseware-builder work is being reviewed on dedicated branches before integration into the default branch.
