---
name: physics-courseware-builder
description: Build rigorous 90-minute high-school physics competition/Strong Foundation course packages from handwritten notes, with formula-aware PPT layout, university-physics QA, exercises, handouts, Xournal boards, and visualizations.
version: 1.0.0-test
---

# Physics Courseware Builder

Use this skill when the user asks to turn handwritten physics notes, Notability pages, PDFs, lesson outlines, or existing physics course material into a complete competition/Strong Foundation teaching package.

Read `references/WORKFLOW.md` before production and use `prompts/MASTER_PROMPT.md` as the orchestration contract.

Core invariants:
1. Teacher notes are the primary source.
2. External sources supplement rather than silently replace.
3. Physics/math QA precedes final layout.
4. Mathematical notation is typeset, not faked with ordinary Unicode.
5. PPT formulas must be compatible with the target environment; avoid SVG as the default if it has produced missing glyphs.
6. Use formula-aware semantic layouts and formula safe zones.
7. Handwritten visual language belongs to annotations, not formal equations.
8. Render and inspect every final slide.
9. Do not release if source, physics, typography, layout, compatibility, pacing, or answer QA fails.
