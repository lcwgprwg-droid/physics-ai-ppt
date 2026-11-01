# M01 Golden Baseline — 向量

Accepted baseline: **M01 V2.0 稳定重制版**

This baseline is the first lesson in the physics competition course sequence and is used to prevent regressions in vector-language teaching, diagram layout, formula sizing, problem-stem completeness and PowerPoint compatibility.

## Lesson contract

- 90 minutes
- 40 slides total
- classroom flow: slides 1–37
- teacher appendix: slides 38–40
- content mother: accepted M01 V1.3 course content
- V2.0 may restructure presentation but must not silently expand or delete the accepted knowledge range

## Required teaching content

- scalar vs vector
- position vector and displacement
- boundary of the statement “vectors can be translated”
- components and unit vectors
- 2D / 3D representation
- vector addition and subtraction
- resultant-magnitude range
- decomposition and projection
- dot product: geometry + coordinates + angle + perpendicularity
- work as a dot product
- cosine-law derivation using vectors
- cross product
- area meaning of cross product
- torque entry and reference-point dependence
- dot vs cross comparison
- motion-vector preview linking M01 to M02/M03

## Accepted visual/compatibility rules

- no giant formula + tiny explanatory text
- formula safe zones
- core vector diagrams use native PowerPoint shapes when practical
- student-facing classroom content stays in the main flow
- source/metadata/QA material goes to teacher appendix
- all classroom exercises have complete stems
- full-deck PDF render and contact-sheet review are mandatory
- missing media relationships = 0
- out-of-bounds objects = 0

## Artifact fingerprints

Binary artifacts are not committed through the text-only connector. Their SHA-256 fingerprints and sizes are recorded in `golden_baseline.json`.

## Validation

Run:

```bash
python tests/validate_m01_golden.py
```

A change that breaks this contract should fail before release.
