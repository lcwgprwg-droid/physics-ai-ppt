---
name: physics-courseware-builder
description: Build rigorous 90-minute high-school physics competition / Strong Foundation course packages from teacher handwritten notes, with 100% note coverage, formula-aware layout, visual-first explanations, complete problem stems, physics QA, and PPT compatibility QA.
version: 1.0.0
---

# Physics Courseware Builder

Use this skill when the user provides handwritten physics notes, Notability/PDF pages, lesson outlines, or an existing physics lesson and wants a complete competition / Strong Foundation teaching package.

Read `references/WORKFLOW.md` before production. Use `prompts/MASTER_PROMPT.md` as the orchestration contract and `references/QUALITY_GATES.md` as the release checklist.

## Non-negotiable invariants

1. **Teacher notes define the lesson boundary.** Unless the teacher explicitly asks to split or merge lessons, one supplied note set is treated as one intended lesson (typically 90 min).
2. **Note coverage must be 100%.** Every formula, proof, worked example, diagram, coordinate-system discussion, derivation, transition, and side note that carries teaching meaning must appear in the coverage matrix and be represented in the course package.
3. **No silent deletion or relocation.** Do not move note content to “next lesson” merely because it seems advanced or dense.
4. **External sources are supplements only.** University physics, Olympiad references, videos, GitHub projects and OER may verify, illustrate, or improve presentation, but must not silently replace the teacher’s structure.
5. **Physics/math QA precedes release.** Check assumptions, signs, units, notation, applicability, initial conditions, proofs, and numerical answers.
6. **Formal math is typeset.** Never fake display equations with ordinary Unicode text.
7. **Formula size is subordinate to teaching hierarchy.** Never use a giant formula merely to fill empty space.
8. **Formula safe zones are mandatory.** No body text, annotations, captions, or diagrams may overlap a formula image/object.
9. **Visual-first rule.** If a concept is naturally geometric or graphical, add a diagram before enlarging the formula. Examples: radians, vector geometry, coordinate systems, secant→tangent, v-t area, polar coordinates, microelement/Riemann sums.
10. **Derivations follow notebook logic.** Prefer `idea → step 1 → step 2 → ... → conclusion`, with one main formula per step.
11. **Problem stems must be complete.** Include physical object/context, known conditions, symbol definitions, ranges/positive direction when needed, and explicit questions. Teacher-note problems are labeled as note-derived; added problems are clearly labeled as supplements.
12. **Handwritten style is annotation only.** Formal equations remain standard mathematical typography; blue handwritten-style notes/arrows/circles may be used as secondary emphasis.
13. **Compatibility is a release gate.** Default PPT formula path is a self-contained high-resolution transparent raster (or another verified compatible format), not SVG when the target environment has shown missing glyphs.
14. **Render every final deck.** Validate media relationships, out-of-bounds objects, missing glyphs, clipping, overlap, and actual PDF/PNG rendering before release.
15. **Do not declare completion if any quality gate fails.**

## Required intermediate artifacts

- `course_spec.json`
- `formula_registry.json`
- `note_coverage_matrix.json` (or CSV)
- `QA_report.md`

## Standard outputs

- Teacher master manuscript `.docx/.pdf`
- Classroom deck `.pptx/.pdf`
- Student handout `.pdf` (optional `.docx`)
- Xournal++ board template `.pdf`
- Interactive visualization `.html` and optional Manim
- Layered exercises `.docx/.pdf`
- `problem_bank.json`
- source review / supplement log

## Stable visual baseline

M02 Golden Test establishes the preferred classroom language:
- 16:9, restrained modern layout;
- graph/diagram + moderate formula + concise explanation;
- no “giant formula + tiny text”;
- multiple formulas are vertically staged, not horizontally squeezed;
- derivation pages use flow + formulas + conclusion;
- course homepage states lesson number, theme, module map and learning goal;
- projector-readable body text and generous whitespace.
