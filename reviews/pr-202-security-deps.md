# Review record: PR #202 — replacing `extract-zip` and `image-size`

**PR:** [#202](https://github.com/singerla/pptx-automizer/pull/202) ·
**Merged:** 2026-08-18 (`3909c16`, amendments in `30aa21a`) ·
**Reviewed with:** Claude Code (the PR itself was also generated with Claude —
see [the last section](#reviewing-an-ai-generated-pr-with-the-same-model) for
why that matters and how we handled it)

This document records what the review found, why the PR was merged despite
its size, and the open questions contributors are invited to think about.

## What the PR does

Removes two dependencies with unpatched high-severity advisories and replaces
them with internal implementations:

- `extract-zip` → `extractToFolder` in `src/helper/jszip-helper.ts`, built on
  jszip (already a core dependency). Rejects absolute and zip-slip entry
  paths, skips symlink entries entirely.
- `image-size` → `src/helper/image-dimensions.ts`, a loop-safe dimension
  parser for PNG, JPEG, GIF, BMP, WebP and SVG.
- Fixes a real pre-existing bug along the way: `compressFolder` returned
  before the output stream finished **and** swallowed write errors, so fs-mode
  callers could clean up the work directory while the zip was still being
  written.

Net dependency count goes down by two; no new dependency is added.

## Why not just wait for upstream fixes?

Both upstreams are confirmed dead ends (verified 2026-08-18):

- **`image-size`**: the GitHub repo was **archived on 2026-06-03** — the
  maintainer wrote they did not want "to deal with the same LLM generated
  'security advisory' about an infinite loop over and over again", with a
  vague plan to revive the project on Codeberg someday. The latest release
  (2.0.2, April 2025) is affected by CVE-2025-71329 / CVE-2025-71330
  (infinite-loop DoS in the ICNS/JXL/HEIF parsers). No patch will appear on
  npm.
- **`extract-zip`**: last published **June 2020**. CVE-2026-19693
  (GHSA-7pqw-9j4j-h8q3, CVSS 8.1) was published 2026-08-17: arbitrary file
  writes outside the extraction directory via a crafted archive containing a
  symlink followed by a same-named regular file. No patched release exists.

Local patching (`yarn patch` / patch-package) would not help either: audit
scanners flag by *version number*, not by code, so every consumer's
`npm audit` and dependabot would stay red regardless. Only removing the
dependency — or an upstream release that will never come — clears the
advisories.

`pptxgenjs` still pins a vulnerable `image-size` transitively; the yarn
resolution to the newest available release remains in place as mitigation.

## How dangerous were these vulnerabilities really?

Honestly: only to deployments that process .pptx files they do not fully
control. Both bugs require a *deliberately crafted* malicious file — no real
PowerPoint output contains duplicate zip entries with symlinks, and no real
slide media accidentally forms an ICNS/JXL header. The extract-zip issue is
narrower still: it only fired in fs mode (`ArchiveFs`); jszip-mode users never
touched that code path.

Two refinements kept us from dismissing them:

1. **"Fully control" is a softer boundary than it feels.** Foreign decks
   arrive through friendly channels — a client's corporate template, an
   agency's design deck, files that passed through third-party tools. For the
   DoS bugs the worst case is a hung worker; for the extract-zip bug it is an
   arbitrary file write on the host, which matters a lot in multi-tenant
   setups.
2. **As a library, the threat model is not ours to assess.** We do not know
   which consumers feed untrusted files into fs mode, and the advisories sit
   in every consumer's CI gate regardless. Removing the dependencies removes
   the reachability analysis for everyone.

An earlier internal evaluation (commit `1c7b728`) had concluded the image-size
DoS alerts were low-risk and could be dismissed. That was correct at the
time; the file-write CVE and the upstream archival both post-date it.

## Review method and findings

The review did not stop at reading the diff:

- Parsers were checked against the format specs (WebP VP8/VP8L/VP8X offsets,
  the JPEG SOF marker set, termination of the JPEG segment walk).
- A 20,000-case fuzz of truncated/corrupted headers produced no unexpected
  exception types; a crafted 2 MB JPEG with 500,000 tiny segments — the
  attack class behind the original advisories — terminates in ~4 ms.
- The zip-slip guards hold for the subtle cases (backslash-separated `..\`
  names on POSIX resolve to a harmless contained filename).
- The full suite (335 tests) passes.

Two real findings came out of it, addressed post-merge in `30aa21a`:

1. **TIFF/ICO dimension detection silently regressed.** `image-size` parsed
   TIFF and ICO; PowerPoint accepts both as media. The internal parser
   throws, and `setRelationTargetCover` falls back to 100×100 defaults — the
   fallback warning also misattributed the failure ("media file not found").
   The warning now reports the real error, and the limitation is documented
   in `image-dimensions.ts`. **Open question:** is a small bounded TIFF IFD
   parser worth adding, or does no downstream consumer use TIFF media?
2. **fs-mode extraction traded streaming for memory.** `extract-zip` streamed
   from disk with roughly constant memory; `extractToFolder` reads the whole
   archive into memory. fs mode exists precisely for memory-bounded
   processing of large decks, so this cuts against the mode's purpose. Fine
   for typical .pptx sizes; documented in the code. **Open question:** should
   this move to a streaming unzip if very large decks meet fs mode in
   practice?

Minor notes (not fixed): `Number(entry.unixPermissions ?? 0)` would mis-parse
an octal *string* (jszip yields numbers on read, so unreachable today), and
`compressFolder` both logs and rethrows errors, reporting them twice.

## Reviewing an AI-generated PR with the same model

This was the project's first externally contributed AI-generated PR — and it
was reviewed by the same model family that wrote it. The obvious objection:
won't a sibling model share the author's blind spots and find nothing new?

Partly, and it is worth being precise about which part:

- **Correlated risk is real for knowledge errors.** If the model believes a
  wrong fact — say, a wrong byte offset in the WebP spec — a same-model
  reviewer plausibly believes it too. No amount of re-reading catches that.
- **Most generation bugs are not knowledge errors.** They are attention
  slips, omissions under context pressure, and spec misreadings — largely
  uncorrelated between runs. A fresh instance reading adversarially, with no
  authorship attachment, catches those at a decent rate.
- **The way out of correlation is not a different model — it is grounding in
  things that are not the model.** The fuzz run, the pathological-JPEG
  timing, and the TIFF probe were adjudicated by execution, not by the
  reviewer's priors. If author and reviewer shared a delusion about JPEG
  parsing, the fuzzer would not share it.
- **Context beats weights.** Both real findings above were not things the
  model "didn't know" — they were things the PR's author *did not have in
  context*: that fs mode exists to bound memory, and that downstream media
  might include TIFF. The reviewer with the most different context (usually
  the maintainer) is often more valuable than a reviewer with different
  weights.

Practical checklist we would apply to the next AI-generated PR:

1. Demand executable evidence; add adversarial tests of your own.
2. Ask "what did this quietly stop doing that used to work?" — diff-vs-
   behavior, not just diff-vs-diff.
3. Spend cross-model second opinions where they pay: domain-knowledge claims
   (binary formats, protocol details), not the overall PR.
4. The maintainer's threat-model questions ("do we really need this?",
   "is it too big?", "why not upstream?") are load-bearing — ask them even
   when, especially when, the patch looks impressive.

There is also a cautionary tale on the other side of this story: the
`image-size` maintainer archived a 2.2k-star package because of low-effort
LLM-generated advisory spam, while this PR used the same technology to
produce a careful, reviewable patch — and its author responded to review
findings with a follow-up commit. The difference between the two outcomes is
verification effort, on both sides. AI-generated contributions are welcome
here on those terms.
