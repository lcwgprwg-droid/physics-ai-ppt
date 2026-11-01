# Physics Courseware Production Workflow v1.0

## Mission
Turn teacher handwritten notes / Notability PDFs into a scientifically rigorous, classroom-ready high-school physics competition / Strong Foundation course package while preserving the teacher’s original lesson logic.

## 0. Lesson-boundary rule
Unless the teacher explicitly asks otherwise, the supplied note set is treated as the intended lesson boundary. If the teacher says the notes are one 90-minute lesson, **do not split content across modules** merely because the material is dense.

## 1. Source priority
1. Teacher handwritten notes / Notability PDF = primary source and lesson skeleton.
2. Teacher-selected verified video/source = review and gap-filling source.
3. Formal Olympiad / Strong Foundation / university-selection physics references.
4. Calculus-based university physics / OER for independent verification.
5. GitHub / open-source projects for visualization, layout, rendering and engineering.

External sources may verify, illustrate, or supplement. They may not silently replace or delete note content.

## 2. Mandatory source reconstruction
Before slide design:
1. Read notes page-by-page.
2. Extract every teaching item: concept, formula, proof, diagram, coordinate-system discussion, example, derivation, side remark, transition, requested exercise.
3. Build `note_coverage_matrix.json` with fields such as:
   - note_page
   - item_id
   - source_text_or_description
   - type
   - planned_slide
   - treatment
   - status
4. Coverage target = **100%**.
5. If any note item is unclear, mark it `needs_review`; do not silently invent a replacement.

## 3. Knowledge reconstruction
Build `course_spec.json` containing:
- lesson number/theme
- target duration
- audience
- prerequisite chain
- lesson map
- timing
- slide semantic types
- notation decisions
- visual requirements
- problem-stem requirements
- note coverage target

The course homepage should state:
- course/competition lesson number
- lesson theme
- 3–5 content modules
- learning goal
- target duration

## 4. Formula pipeline
Handwriting/screenshot -> OCR candidate LaTeX -> compare with original notes -> physics/math QA -> canonical LaTeX/Typst -> render -> presentation.

### Typography rules
- Variables italic: `t, x, v, a, r, theta...`
- Differential operator upright: `\mathrm{d}t`, `\mathrm{d}x`.
- Graph axes use mathematical typography.
- Inline formulas align with body-text baseline and visual scale.
- Keep canonical source in `formula_registry.json`.
- No fake Unicode display fractions/superscripts.
- No SVG as the default when target playback has shown missing glyphs.

### Formula hierarchy
- Inline: about body-text visual height.
- Derivation: about 1.2–1.5x body visual height.
- Core conclusion: generally <= 1.7x body visual height.
- Formula size must never be increased simply to fill space.
- If a formula becomes the dominant visual without pedagogical reason, shrink it and add a diagram / explanation / whitespace instead.

### Formula safe zone
- Formula bounding box is a protected area.
- No body text, caption, handwritten annotation, image or other equation may overlap it.
- Multiple formulas on the same slide should normally be vertically staged.

## 5. Visual-first rule
When the content can be represented geometrically or graphically, add the visual before enlarging formulas.

Typical triggers:
- radian measure -> circle/arc/radius/angle diagram
- vector products -> parallelogram / projection diagram
- Cartesian/polar/cylindrical/spherical coordinates -> coordinate-system diagrams
- average to instantaneous velocity -> secant/tangent graph
- derivative -> tangent/local slope
- integral -> Riemann rectangles / signed area
- microelement method -> sliced process diagram
- displacement vs distance -> positive/negative area diagram

## 6. Semantic slide layouts
Classify every slide before placing objects:
- cover / lesson-map
- concept
- statement
- two-column
- three-card
- graph + formula
- diagram + formula
- derivation
- relation-map
- problem-stem
- worked-example
- visualization
- summary / exit-ticket

Do not begin with arbitrary x/y coordinates. Choose a semantic component first.

## 7. Derivation-page rule
Derivations should look like a clear notebook argument rather than a wall of equations.

Preferred structure:
1. Method/idea strip
2. Step 1 formula + one-line meaning
3. Step 2 formula + one-line meaning
4. ...
5. Final conclusion strip

Rules:
- one main formula per step;
- preserve the teacher-note derivation order;
- do not skip an intermediate step that appears in the notes;
- if there are multiple formulas, reduce font size instead of overlapping them.

## 8. Problem-stem rule
Every exercise/problem slide must be understandable without reading the solution slide.

Required elements when applicable:
- physical object/context
- known conditions
- symbol definitions
- time/spatial range
- positive direction / coordinate convention
- explicit target questions

Original-note problems: label as note-derived.
Added problems: label as classroom supplement / adapted original; do not impersonate a source problem.

Preferred classroom sequence:
`complete stem -> prediction/route -> derivation/solution -> self-check`.

## 9. Physics QA
Independently check:
- assumptions and applicability
- sign conventions
- units/dimensions
- direction/vector meaning
- boundary/initial conditions
- proof logic
- numerical answers
- graph interpretation

For calculus-based kinematics:
- `v = dx/dt`
- `a = dv/dt = d2x/dt2`
- finite increment: `Δx ≈ v(t)Δt`
- differential relation: `dx = v(t)dt`
- displacement = signed v-t area
- distance requires `|v|` when velocity changes sign
- `a < 0` does not automatically mean slowing down
- integration constants encode physical initial state

## 10. PPT compatibility rules
Before release:
1. Use self-contained media.
2. Do not clone slide XML without correctly copying media relationships.
3. Inspect PPT package for missing media targets.
4. Check shape bounds / out-of-slide objects.
5. Render the full deck to PDF/PNG.
6. Visually inspect representative and high-risk pages.
7. Reject any page with missing image, `?`, clipping, overlap, or font substitution that changes meaning.

## 11. Exercise system
Default homework: 5–10 questions, typically 8.
Suggested composition:
- 2 basic
- 2 method
- 2 Olympiad / Strong Foundation
- 1 comprehensive
- 1 calculus/university bridge

Recalculate numerical answers independently. For uncertain copyrighted sources, create an equivalent/adapted problem and label it clearly.

## 12. Quality gates
A. Note coverage = 100%
B. Source fidelity
C. Physics/math correctness
D. Formula typography
E. Visual/layout readability
F. Complete problem stems
G. PPT compatibility/rendering
H. Classroom pacing
I. Exercise-answer verification

Release only if all mandatory gates pass. See `QUALITY_GATES.md`.

## 13. Standard outputs
- teacher manuscript `.docx/.pdf`
- classroom deck `.pptx/.pdf`
- student handout `.pdf`
- Xournal++ board template `.pdf`
- interactive HTML / optional Manim
- layered exercise set `.docx/.pdf`
- `problem_bank.json`
- `course_spec.json`
- `formula_registry.json`
- `note_coverage_matrix.json`
- `QA_report.md`

## 14. Golden-test policy
M02 is the first stable Golden Test. Its validated lessons are reusable constraints, not topic-specific decorations:
- preserve all note content;
- visual-first explanation;
- restrained formula sizing;
- formula safe zones;
- notebook-style derivations;
- complete stems;
- coordinate-system diagrams when notes discuss coordinates;
- actual compatibility rendering before release.

A new renderer/skill revision should not be called stable until it reproduces M02-level quality and then passes a second, substantially different course (e.g. M01 vectors or M03 microelement method).
