---
name: physics-courseware-builder
description: Build rigorous 90-minute high-school physics competition / Strong Foundation course packages from teacher handwritten notes, with 100% note coverage, formula-aware layout, visual-first explanations, complete problem stems, physics QA, and PPT compatibility QA.
version: 1.1.0
---

# Physics Courseware Builder

Use this skill when the user provides handwritten physics notes, Notability/PDF pages, lesson outlines, or an existing physics lesson and wants a complete competition / Strong Foundation teaching package.

Read `references/WORKFLOW.md` before production. Use `prompts/MASTER_PROMPT.md` as the orchestration contract, `references/QUALITY_GATES.md` as the release checklist, and `USAGE.md` for invocation examples.

## Quick invocation

After uploading the source notes, the teacher can simply say:

> 按 Physics Courseware Builder 标准生成下一节90分钟竞赛课，严格按我的笔记，不遗漏，不擅自补题干，按 M01–M04 稳定版标准完成 PPT、PDF、覆盖矩阵和全页 QA。

Optional shorter trigger:

> 按竞赛课件 Skill 生成下一节。

If the teacher says the shorter trigger, infer the same production contract from this skill unless they explicitly override it.

## Non-negotiable invariants

1. **Teacher notes define the lesson boundary.** Unless the teacher explicitly asks to split or merge lessons, one supplied note set is treated as one intended lesson, typically 90 minutes.
2. **Note coverage must be 100%.** Every formula, proof, worked example, diagram, coordinate-system discussion, derivation, transition, and side note that carries teaching meaning must appear in the coverage matrix and be represented in the course package.
3. **No silent deletion or relocation.** Do not move note content to “next lesson” merely because it seems advanced or dense.
4. **Do not invent missing source conditions.** If a handwritten example contains only a diagram/derivation and no complete text stem, preserve it as a diagram-derived example and explicitly mark the source gap. Never fabricate a plausible stem and present it as the teacher’s original problem.
5. **External sources are supplements only.** University physics, Olympiad references, videos, GitHub projects and OER may verify, illustrate, or improve presentation, but must not silently replace the teacher’s structure.
6. **Physics/math QA precedes release.** Check assumptions, signs, units, notation, applicability, initial conditions, proofs, and numerical answers.
7. **Formal math is typeset.** Never fake display equations with ordinary Unicode text.
8. **Formula size is subordinate to teaching hierarchy.** Never use a giant formula merely to fill empty space.
9. **Formula safe zones are mandatory.** No body text, annotations, captions, or diagrams may overlap a formula image/object.
10. **Visual-first rule.** If a concept is naturally geometric or graphical, add a diagram before enlarging the formula.
11. **Derivations follow notebook logic.** Prefer `idea → step 1 → step 2 → ... → conclusion`, with one main formula per step.
12. **Problem stems must be complete when the source is complete.** Include physical object/context, known conditions, symbol definitions, ranges/positive direction when needed, and explicit questions. Teacher-note problems are labeled as note-derived; added problems are clearly labeled as supplements.
13. **Handwritten style is annotation only.** Formal equations remain standard mathematical typography; handwritten-style notes/arrows/circles may be used as secondary emphasis.
14. **Compatibility is a release gate.** Default PPT formula path is a self-contained high-resolution transparent raster or another verified compatible format, not SVG when the target environment has shown missing glyphs.
15. **Prefer native PowerPoint shapes for fragile core diagrams.** Rebuild simple geometry/axes/arrows with native shapes when transparent-image cropping or platform compatibility is a risk.
16. **Render every final deck.** Validate media relationships, out-of-bounds objects, missing glyphs, clipping, overlap, and actual PDF/PNG rendering before release.
17. **Do not declare completion if any quality gate fails.**

## Required intermediate artifacts

- `course_spec.json`
- `formula_registry.json`
- `note_coverage_matrix.json` or CSV
- `QA_report.md`

## Standard outputs

Minimum classroom release:
- Classroom deck `.pptx`
- Rendered deck `.pdf`
- Note coverage matrix
- Full-deck contact sheet / visual QA image
- QA report

Full course package when requested:
- Teacher master manuscript `.docx/.pdf`
- Student handout `.pdf` and optional `.docx`
- Xournal++ board template `.pdf`
- Interactive visualization `.html` and optional Manim
- Layered exercises `.docx/.pdf`
- `problem_bank.json`
- source review / supplement log

## Stable visual baseline

The accepted M01–M03 baselines and M04 production candidate establish the classroom language:
- 16:9, restrained modern layout;
- graph/diagram + moderate formula + concise explanation;
- no “giant formula + tiny text”;
- multiple formulas are vertically staged, not horizontally squeezed;
- derivation pages use flow + formulas + conclusion;
- course homepage states lesson number, theme, module map and learning goal;
- projector-readable body text and generous whitespace;
- source-derived incomplete examples are marked, not fabricated;
- high-risk diagrams should use native PowerPoint shapes when possible;
- full-deck render and contact-sheet review are mandatory.

## Regression expectation

Before changing layout/rendering rules, preserve the accepted Golden Baselines:
- M01 — vectors / native geometry / student-facing structure
- M02 — formula typography / calculus / complete stems
- M03 — microelement geometry / six-example derivations / compatibility

M04 is the next production candidate and stresses:
- relative motion notation;
- moving reference frames;
- wind/boat/person vector problems;
- associated-motion constraints;
- strict distinction between reference-frame relations and geometric constraints.
