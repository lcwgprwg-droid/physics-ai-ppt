# Physics Courseware Production Workflow v1.0-test

## Mission
Turn teacher handwritten notes / Notability PDFs into a scientifically rigorous, classroom-ready 90-minute high-school physics competition + Strong Foundation course package.

## Source priority
1. Teacher handwritten notes / Notability PDF.
2. Verified teacher-selected video.
3. Formal high-school physics Olympiad framework.
4. Strong Foundation / university selection / CPhO / IPhO / calculus-based university physics.
5. GitHub / OER for visualization, layout, engineering and independent QA.

Never silently replace the teacher's intended structure.

## Formula pipeline
Handwriting/screenshot -> OCR candidate LaTeX -> compare with original notes -> physics/math QA -> canonical LaTeX/Typst -> render -> PPT.

Rules:
- No SVG as default when target playback has shown missing glyphs.
- No fake Unicode display equations.
- Keep canonical source in formula_registry.json.
- Variables italic; differential operator upright.
- Graph axes use math typography.
- Inline formulas must visually match surrounding prose.
- For M02 classroom notation, prefer t consistently.

## Formula hierarchy
- Inline: about body-text visual height.
- Derivation: about 1.3–1.6x body.
- Core conclusion: up to about 1.8–2.0x body.
Never make every formula huge.

## Semantic layouts
concept / statement / two-column / three-card / graph+formula / derivation / relation-map / example / visualization / summary.

## Layout rules
- 16:9.
- One visual center per slide.
- No “giant formula + tiny explanatory text”.
- Formula safe zones are mandatory.
- Handwritten styling is annotations only.
- Examples use problem -> route -> reasoning -> self-check.

## Physics QA
Check assumptions, signs, units, applicability, initial conditions and numerical answers.
For calculus-based kinematics:
- v = dx/dt
- a = dv/dt = d2x/dt2
- Δx ≈ v(t)Δt for a short finite interval
- dx = v(t)dt as differential relation
- displacement = signed v-t area
- distance requires |v| if sign changes
- a<0 does not automatically mean slowing down

## QA gates
A Source fidelity
B Physics correctness
C Formula typography
D Layout/readability
E PPT compatibility
F Classroom pacing
G Exercise answer verification

Release only if all gates pass.
