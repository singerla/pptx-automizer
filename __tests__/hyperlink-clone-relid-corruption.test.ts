import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import Automizer from '../src/index';

/**
 * Regression test for a relationship-id corruption bug: `Shape.appendToSlideTree()`
 * and `Shape.modifySlideTree()` used to unconditionally overwrite a Hyperlink
 * shape's `r:id` with a separately precomputed id whenever the shape already
 * carried a resolved hyperlink (the `slide.modifyElement()`/`addElement()` path
 * for a template shape addressed by name/selector, classified as
 * `ElementType.Hyperlink`). That precomputed id had no matching `<Relationship>`
 * entry, leaving an `r:id` with nothing to resolve to, which made PowerPoint
 * report "found a problem with content" and offer to repair (silently
 * dropping the shape).
 *
 * Both scenarios below write via the normal `Automizer.write()`, which is
 * wrapped by `__tests__/helpers/setup-pptx-invariants.ts` to fail on any
 * `r:id`/`r:embed` reference in the written archive that has no matching
 * relationship — the assertions here are an explicit, self-documenting
 * check of the same property for this specific shape.
 */

async function readSlidePart(
  outputPath: string,
  slideNumber: number,
): Promise<{ slideXml: string; relsXml: string }> {
  const zip = await JSZip.loadAsync(fs.readFileSync(outputPath));
  const slideXml = await zip
    .file(`ppt/slides/slide${slideNumber}.xml`)!
    .async('text');
  const relsXml = await zip
    .file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)!
    .async('text');
  return { slideXml, relsXml };
}

function declaredRelationshipIds(relsXml: string): string[] {
  return Array.from(relsXml.matchAll(/<Relationship\s[^>]*Id="([^"]+)"/g)).map(
    (match) => match[1],
  );
}

function assertNoDanglingRelIds(slideXml: string, relsXml: string): void {
  const declaredRIds = declaredRelationshipIds(relsXml);
  const usedRIds = Array.from(
    slideXml.matchAll(/r:(?:id|embed)="([^"]+)"/g),
  ).map((match) => match[1]);

  expect(usedRIds.length).toBeGreaterThan(0);
  usedRIds.forEach((rId) => expect(declaredRIds).toContain(rId));
}

test('modifyElement on a cloned hyperlink shape does not corrupt its relationship id', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `hyperlink-clone-relid-corruption-modify.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  const result = await pres
    .addSlide('link', 1, (slide) => {
      // No-op modification: this only re-positions the "LinkToSlide" shape,
      // whose relationship is already set up, exercising the redundant
      // hyperlink processing in Shape.modifySlideTree() without a callback
      // that would mask corruption by discarding the hyperlink altogether.
      slide.modifyElement('LinkToSlide', []);
    })
    .write(outputFile);

  expect(result.slides).toBe(2);

  const { slideXml, relsXml } = await readSlidePart(outputPath, 2);
  assertNoDanglingRelIds(slideXml, relsXml);

  const rIdMatch = slideXml.match(
    /<p:cNvPr id="\d+" name="LinkToSlide"[^>]*>[\s\S]*?<a:hlinkClick r:id="([^"]+)"/,
  );
  expect(rIdMatch).not.toBeNull();
  const linkRid = rIdMatch![1];

  const relPattern = new RegExp(
    `<Relationship\\s[^>]*Id="${linkRid}"[^>]*Target="slide2\\.xml"`,
  );
  expect(relsXml).toMatch(relPattern);
});

test('addElement of a template hyperlink shape creates its relationship', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `hyperlink-clone-relid-corruption-append.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('link', 1, 'LinkToSlide');
    })
    .write(outputFile);

  expect(result.slides).toBe(2);

  const { slideXml, relsXml } = await readSlidePart(outputPath, 2);
  assertNoDanglingRelIds(slideXml, relsXml);

  const rIdMatch = slideXml.match(/<a:hlinkClick r:id="([^"]+)"/);
  expect(rIdMatch).not.toBeNull();
  const linkRid = rIdMatch![1];

  const relPattern = new RegExp(
    `<Relationship\\s[^>]*Id="${linkRid}"[^>]*Target="(?:\\.\\./slides/)?slide2\\.xml"`,
  );
  expect(relsXml).toMatch(relPattern);
});
