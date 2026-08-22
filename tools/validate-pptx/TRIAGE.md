# Triage recipe: a generated .pptx makes PowerPoint ask to repair

A generic, step-by-step way to find out *why* a produced deck triggers the
"PowerPoint found a problem with content" repair prompt, and to fix both the
file and the code path that broke it. No PowerPoint installation is needed
until the final verification step.

## 1. Validate the file

Build and run the Tier-2 validator (Open XML SDK) against the deck directly —
skip the allowlist so *every* finding is visible:

```bash
docker build -q -t pptx-validator tools/validate-pptx
docker run --rm -v /path/to/dir:/work:ro pptx-validator --verbose /work/broken.pptx
```

The validator is the genuine oracle for the repair-prompt bug class: it flags
package-level relationship violations and schema child-order violations that
XML assertions and pixel diffs both miss. If the SDK cannot even open the file
(`[OpenFailed]`), that is the strongest possible signal (usually dangling
relationships).

## 2. Cluster the errors into classes

Do not read 50+ errors one by one. Group them — in practice they collapse into
a handful of classes. The classes seen so far, most severe first:

| Class | Validator message shape | Meaning |
| --- | --- | --- |
| Wrong-part relationship | `'PresentationPart{...}' cannot have a relationship that targets part 'SlideLayoutPart{...}'` | A rel was registered on a part that may not own it (layouts belong to their slideMaster, not to `presentation.xml`). |
| Duplicate singleton rel | `can only have one instance of relationship that targets part '.../slideLayout'` (or `notesSlide`) | A slide got a *second* layout/notes rel — typically a rel copied wholesale from a source slide. |
| Schema child order | `The element has unexpected child element '...:solidFill'` (or `pPr`, `bodyPr`, `ln`, …) | Children of an element are present but in the wrong sequence, or duplicated (`a:pPr` must be unique and first in `a:p`; `a:rPr` children follow CT_TextCharacterProperties order). |
| Unknown attribute | `The 'xyz' attribute is not declared` | Usually pre-existing noise carried in from the source templates. Check whether the *input* templates already show it (step 3) before blaming the pipeline. |

## 3. Separate inherited noise from pipeline damage

Run the validator on the **input templates** too. Anything that already exists
there is inherited noise (real-world templates always carry some) and is
almost never the repair trigger — PowerPoint opened those inputs fine.
Everything that appears **only in the output** was produced by the pipeline
and is a suspect.

## 4. Read the raw XML of the suspects

```bash
mkdir x && cd x && unzip -q ../broken.pptx
```

Look at the flagged parts (`ppt/_rels/presentation.xml.rels`, the slide's
`_rels/*.rels`, the slide XML at the reported XPath). Two forensic markers help
attribute each defect:

- **Relationship-id suffixes** — pptx-automizer marks every rel it adds with an
  `-created` suffix (`rId7-created`). Original template rels have plain ids.
  A *broken* `-created` rel means the library (or its caller) wrote it.
- **Target numbering** — a rel target numbered in the *source* template's
  scheme (e.g. `slideLayout1.xml` on a slide whose real layout is
  `slideLayout41.xml`) means a rel was copied verbatim without remapping.
- **Inline `xmlns:` declarations** on single elements (e.g.
  `<a:hlinkClick xmlns:r="…">`) mean the element was programmatically inserted
  (xmldom serializes it that way), not copied from a template.

Also check whether the slide XML *references* the bogus rels (search the slide
for the rel id). An `a:hlinkClick` pointing at a slideLayout or notesSlide rel
means a hyperlink rel-id was resolved against the wrong relationships and the
copy step cloned whatever rel happened to sit at that id.

## 5. Bisect to the actual repair trigger (optional)

PowerPoint tolerates some violations and repairs on others. If you need to
know *which* class triggers the prompt, fix one class at a time into separate
copies (`fix-A.pptx`, `fix-B.pptx`, …) and open each in real PowerPoint.
A small Python/lxml script over the unzipped parts is enough; re-zip with
`zipfile` (store the same member names, deflate).

Typical fixes per class:

- **Wrong-part rel**: delete the rel from the offending `.rels`; confirm the
  target part is still reachable from its correct owner (e.g. each layout from
  a master's `.rels`).
- **Duplicate singleton rel**: keep the first/original rel, drop the copies,
  then remove or re-point any element still referencing a dropped id —
  a dangling `r:id` is itself a repair trigger.
- **Child order**: sort children per the schema sequence (for `a:rPr` see
  `RPR_CHILD_ORDER` in `src/helper/modify-text-helper.ts`); drop duplicate
  `a:pPr` beyond the first in each `a:p`.

Re-run the validator after each fix; sanity-check the result opens by
converting it headlessly (`soffice --headless --convert-to pdf fix.pptx`),
then have the final candidate opened in real PowerPoint.

## 6. Fix the code path, not just the file

For each output-only class, find the writer in the library and add the missing
guard, then remove the matching allowlist entry (see the rule at the top of
`allowlist.json`) and add a Tier-0 assertion so the class cannot silently
return. Known writer locations for the classes above:

- rels appended to `ppt/_rels/presentation.xml.rels`:
  `src/classes/content-type-registry.ts` (`addToPresentation`).
- rels copied from a source slide for hyperlinks:
  `src/helper/hyperlink-processor.ts` (`copyMultipleHyperlinks`) — must only
  copy rels whose `Type` is actually a hyperlink.
- `a:rPr` style setters: `src/helper/modify-text-helper.ts`
  (`sortChildrenBySchema` after mutation).
- duplicate `a:pPr` mid-paragraph (`[pPr, r, pPr, r, endParaRPr]`): upstream
  pptxgenjs — `genXmlTextBody` emits `genXmlParagraphProperties` per text run
  inside one `<a:p>`, so multi-run lines built via the `generate()` bridge
  carry one `a:pPr` per run. No caller-side workaround exists: the
  `!options.bullet` fallthrough (ISSUE#589) writes a populated
  `indent="0" marL="0"` + `<a:buNone/>` pPr for every non-bullet run, so only
  empty pPr blocks get stripped. Not automizer XML; possible defensive fix is
  a bridge-side sanitizer dropping 2nd+ `a:pPr` per `a:p` (the duplicates are
  byte-identical); otherwise track under the allowlist's UPSTREAM section and
  revisit on pptxgenjs upgrades.

Finally, check which library version produced the broken deck — the defect may
already be fixed on a newer release and the consumer only needs to upgrade.
