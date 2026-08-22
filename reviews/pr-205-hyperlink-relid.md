# Review record: PR #205 — hyperlink relationship-id corruption on clone

**PR:** [#205](https://github.com/singerla/pptx-automizer/pull/205) ·
**Resolution:** superseded by the maintainer branch
`fix/hyperlink-rel-type-guards` — the PR's commit was cherry-picked with the
contributor's authorship preserved, then corrected and extended ·
**Reviewed with:** Claude Code (Fable 5); the PR discloses it was written
with Claude Sonnet 5 — see [the last section](#reviewing-across-the-model-power-gradient)
for why that pairing turned out to be the most instructive part.

This document records what the review confirmed, the false assumption it
uncovered, the real-world corruption class the PR could not have known about,
and why we fixed the PR in-house instead of sending it back.

## What the PR claims and does

Cloning or modifying a template shape that already carries a hyperlink could
corrupt the output: PowerPoint reported "found a problem with content" and
silently dropped the shape on repair. The PR's diagnosis: both
`Shape.appendToSlideTree()` and `Shape.modifySlideTree()` overwrite the
shape's `<a:hlinkClick r:id>` with a precomputed id that has no matching
`<Relationship>` entry. It removes the overwrite from **both** paths, reworks
the guard in `ModifyHyperlinkHelper.addHyperlink()` to check whether an
existing `r:id` resolves before skipping, and adds a regression test for each
path. All 346 tests pass on the branch, and the description says the new
tests fail against pre-fix code.

## What the review confirmed

Half of it, completely. The **modify path** is broken exactly as described:
`modifySlideTree()` runs the overwrite *after* `editTargetHyperlinkRel()` has
verified the existing relationship is already correct, leaving a dangling
`r:id` in the common no-op/reposition case. The PR's modify test fails
against `main` and passes with the fix. That diagnosis, the removal, and the
test were merged unchanged.

## The false assumption

The claim that the **append path** produced dangling ids too does not
survive re-derivation. On `main`, `addHyperlink()` creates its relationship
*before* the "link already set" guard returns — and the id it mints
(`rId${getMaxId(...)+1}`) always coincides with the precomputed `createdRid`,
because both are max+1 over the same rels file and nothing writes to it in
between. The append path worked, by an unplanned but deterministic
coincidence. The tell was checkable in one run: **the PR's own append test
passes against pre-fix `main`.** Only the modify test fails there. The PR's
causal story — "the guard saw the unmatched id and skipped" — describes code
that does not exist; the guard runs after the relationship is created.

Worse, removing the append-path rewrite introduced a regression. Without it,
a cloned shape keeps its *source* rId, and the new "existing `r:id` resolves
→ skip" guard false-positives whenever that id collides with an unrelated
relationship already declared on the target slide — likely in practice,
since every slide's rels start at rId1 (layout, notes, images…). Repro with
fixtures already in the repo: `addElement` of the `LinkToSlide` shape
(source `r:id="rId3"`) onto `SlideWithImages` slide 2, whose `rId3` is an
image relationship. On the PR branch the hyperlink silently resolves to the
image rel — an *internally consistent mislink* that no dangling-id check can
see. The PR's fixtures dodge it only because its target slide's rels contain
nothing beyond `rId1`.

Note what did **not** go wrong: the code is clean, the tests are real, the
claim "tests fail against pre-fix code" was half-true and honestly meant.
The failure is a plausible, coherent, wrong causal narrative — and test
fixtures unconsciously shaped so they confirm it. When one agent writes the
code, the story, and the evidence, all three can share a blind spot. The
review found the bug not by reading harder but by refusing the narrative:
re-derive the mechanism, then execute a counterexample the author didn't
choose.

## What shipped instead

- The `r:id` rewrite stays on the append path — rewriting to the freshly
  reserved id is precisely what makes a cloned shape collision-proof — and
  is removed only from the modify path, the actual bug.
- The PR's guard idea is kept but made **Type-aware**: skip only when the
  existing `r:id` resolves to a hyperlink/slide relationship; create the
  missing relationship for an unmatched id; allocate a fresh id when the id
  resolves to a relationship of the wrong Type.
- Regression tests for the collision scenario, alongside the PR's tests.

## The real-world extension the PR could not have known

The day after the review, a parallel forensic session triaged a
field-generated deck (produced by a downstream pipeline pinning an older
release, anonymized here) that triggers the repair prompt. Same corruption
class, different door: `HyperlinkProcessor.copyMultipleHyperlinks()` — the
GenericShape import path for multi-hyperlink shapes, untouched by the PR —
resolves each `a:hlinkClick r:id` against the *source* slide's rels and
cloned whatever relationship sat at that id, Type and all. A stale id at the
classic rId1/rId2 positions cloned the slideLayout or notesSlide
relationship verbatim, complete with source-numbered targets. The Open XML
SDK calls it precisely: *"can only have one instance of relationship that
targets part"* — an OPC singleton violation, and this class survives every
internal consistency check because nothing dangles.

The maintainer branch therefore also Type-filters that copy (only
hyperlink/slide rels), strips hyperlinks it cannot wire up (a dangling
`r:id` is itself a repair trigger), and teaches the test suite's package
invariants two new error classes: duplicate slideLayout/notesSlide
relationships, and any `hlinkClick`/`hlinkHover` resolving to a structural
relationship. A patched-fixture regression test reproduces the field
corruption and fails against pre-fix code; the fixed output validates with
zero errors under the SDK.

## "Please correct this PR" versus fixing it here

We did not meter tokens precisely, but the ratios are clear enough to
reason with. The adversarial review — re-deriving the id arithmetic, running
the PR's tests against pre-fix code, building the collision counterexample —
was the expensive part. Writing the corrected fix afterwards cost maybe half
again as much, because everything the fix needed was already in the
reviewer's context: the mechanism, the counterexample, the fixture layout,
the field-deck forensics.

A request-changes round-trip inverts that economy. The contributor's agent
would rebuild context from a prose review, regenerate, and the maintainer
would then owe a *full* re-review of v2 — adversarial again, since the
first version's tests had already demonstrated they could green-light a
regression. That is the review cost a second time, plus days of latency,
plus a hard limit nobody can route around: the field-deck evidence lives in
private material that cannot leave the maintainer's side, so the
contributor's v2 could never have covered `copyMultipleHyperlinks` at all.
For a change this size, review-then-fix-in-house — with the contribution
cherry-picked so authorship survives, and a reply that shows the
counterexample instead of asserting it — is cheaper, faster, and ends in a
strictly better patch. For a large PR the calculus can flip; the deciding
variable is whether the review already left the fix's full context loaded.

## Reviewing across the model-power gradient

The PR #202 record asked whether a model can review a PR written by the
same model family, and answered: mostly yes, if the review is grounded in
execution rather than opinion. This PR poses the harder follow-up question,
because here the pairing was asymmetric — authored with a mid-tier model,
reviewed with a stronger one — and the gap showed up nowhere you would
expect. Not in the code, which was idiomatic. Not in rigor's outward forms:
tests written, suite green, claims stated, model disclosed. The gap was one
level up, in the *causal story* — a mechanism explained confidently,
symmetrically, and wrongly, with evidence that had quietly arranged itself
to agree.

That is worth sitting with, because it predicts what capability drift will
look like in open-source at large. Contributions from weaker models will
not usually fail loudly; loud failures get filtered before the PR is
opened. What survives the author's own checks and arrives in your review
queue is precisely the *plausible* failure — and as the gradient between
authoring models widens, the review queue becomes a spectrum of
convincing-but-unverified narratives, most of them true. Reading harder
does not sort them. Two things do. First, re-derive the mechanism from the
code as if the description did not exist — the description is the one part
of a PR its author's blind spot is guaranteed to have written. Second,
bring inputs the author did not choose: other fixtures, other code paths,
a validator, a broken deck from the field. An author who controls both the
code and its evidence controls the verdict; the reviewer's entire leverage
is everything the author never saw.

And the gradient points at the reviewer too. The stronger model's review
was not smarter reading — it was a refusal to accept a story plus one
executed counterexample, habits that are model-independent and that today's
strongest model had better keep, because it is tomorrow's mid-tier author.
The contributor here did the three things that made the asymmetry
productive rather than adversarial: disclosed the model, kept the claims
separable from the code, and made the causal story explicit enough to be
checked. That is why the wrong half was cheap to find and the right half
was worth keeping. Write your PRs — and your reviews — for an audit by
something stronger than you, arriving with context you don't have. On
current trends, that is not a hypothetical reviewer. It is just the next
release.
