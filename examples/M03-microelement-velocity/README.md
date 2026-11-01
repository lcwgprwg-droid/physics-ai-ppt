# M03 Golden Test — 微元法求速度

M03 is the second Golden Test for the Physics Courseware Builder.

## Accepted stable baseline

- Version: **V1.1**
- Lesson: **90 min**
- Slides: **38**
- Source boundary: the original 3-page handwritten note set is one lesson and must not be silently split.
- Required source examples: **6 / 6**
- Required note coverage items: **14 / 14**
- Missing PowerPoint media relationships: **0**
- Out-of-bounds objects: **0**
- LibreOffice/PDF render QA: **passed**
- Full-deck contact-sheet review: **passed**

## What this Golden Test protects

1. **Source fidelity**
   - The notes define the lesson boundary.
   - Examples 1 and 2 have incomplete prose stems in the source notes; generated materials must not invent unsupported conditions.

2. **Microelement-method teaching logic**
   - Δt → 0
   - draw before/after positions
   - convert motion to micro-displacements such as `v Δt`
   - use geometry / similarity / projection / trigonometry
   - divide by Δt and take the instantaneous limit

3. **Required content**
   - planar and 3D velocity vectors
   - unit-circle trigonometry
   - small-angle relation `sin θ ≈ θ ≈ tan θ`
   - all 6 source examples
   - sine law with circumcircle form `= 2R`
   - cosine law
   - final intersection-C problem with the cyclic/circumcircle bridge

4. **Visual and compatibility rules**
   - diagrams before giant formulas
   - formulas must stay inside safe zones
   - fragile transparent diagrams should be converted to native PowerPoint geometry
   - accepted native-shape compatibility pages: **1, 7, 9**
   - the deck must render cleanly before release

## Baseline artifact contract

Binary artifacts are not committed through the text-only GitHub connector. Their accepted filenames, SHA-256 hashes, sizes, and QA status are stored in `golden_baseline.json`.

Any future renderer change that materially changes M03 should be compared against this baseline and must still pass `tests/validate_m03_golden.py`.
