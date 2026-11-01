# Physics Courseware Builder — Quality Gates v1.0

A course package is releasable only when all mandatory gates pass.

## Gate A — Note coverage
Mandatory.
- Coverage matrix exists.
- Every meaningful note item is mapped to a slide/artifact.
- Coverage = 100%.
- Proofs, examples, diagrams, coordinate discussions and side derivations count as meaningful content.
- No content is silently moved to another lesson.

FAIL examples:
- sine law present in notes but absent from PPT;
- coordinate-system discussion omitted;
- derivation reduced to final formula only when notes contain intermediate steps.

## Gate B — Source fidelity
Mandatory.
- Teacher-note material and external supplements are distinguishable.
- Teacher ordering/logic is preserved unless explicitly redesigned with the teacher’s approval.
- Unclear note content is flagged rather than invented.

## Gate C — Physics/math correctness
Mandatory.
Check:
- assumptions / domain of validity;
- units/dimensions;
- vector direction/sign conventions;
- initial/boundary conditions;
- proof logic;
- numerical results;
- graph interpretation.

## Gate D — Formula typography
Mandatory.
- formal equations use a math engine;
- variables italic, differential operator upright;
- graph axes are math typeset;
- no fake Unicode display fractions;
- inline math matches body scale;
- formula registry preserves canonical source.

## Gate E — Formula/layout balance
Mandatory.
- no giant formula + tiny explanatory text;
- formula safe zones are respected;
- no formula overlaps text/image/another formula;
- multiple formulas are staged vertically or split across slides;
- visual hierarchy reflects teaching importance, not equation length.

## Gate F — Visual explanation
Mandatory when the concept is naturally visual.
Expected examples:
- radians -> arc/radius/angle diagram;
- coordinates -> coordinate-system diagram;
- vectors -> projection/parallelogram;
- secant/tangent -> graph;
- integral -> area/rectangles;
- microelement method -> sliced-process diagram.

## Gate G — Derivation quality
Mandatory for note derivations/proofs.
- method/idea shown;
- intermediate steps preserved;
- one main formula per step when possible;
- conclusion clearly separated;
- derivation resembles a readable board/notebook argument.

## Gate H — Problem-stem completeness
Mandatory for exercises/examples.
A problem must include, when applicable:
- physical object/context;
- known conditions;
- symbol definitions;
- interval/range;
- positive direction / coordinate convention;
- explicit questions.

A slide containing only a formula plus “求…” fails this gate unless the source itself intentionally defines a micro-exercise that way.

## Gate I — PPT compatibility
Mandatory.
- no missing media relationships;
- no out-of-bounds objects;
- no missing glyph/image placeholders;
- full deck renders successfully to PDF/PNG;
- high-risk pages visually inspected;
- avoid SVG as default when target environment has shown incompatibility.

## Gate J — Classroom pacing
Mandatory.
- deck density is plausible for the target duration;
- the teacher can actually explain the derivations/examples in the allotted time;
- dense note content is reorganized, not deleted.

## Gate K — Exercise verification
Mandatory when exercises are included.
- numerical answers independently recalculated;
- answer and solution route checked;
- source/copyright status handled appropriately;
- added questions clearly labeled as supplement/adaptation.

## Release status
Use one of:
- `FAIL` — one or more mandatory gates fail.
- `PARTIAL PASS` — core logic works but one or more release gates are not yet fully tested.
- `PASS` — all mandatory gates pass.
- `GOLDEN` — PASS + accepted by teacher as a reusable visual/teaching baseline.

M02 V1.6.2 is the first accepted visual/content baseline for the Golden Test rules; future versions may supersede it while preserving the invariants above.
