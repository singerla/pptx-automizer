import * as fs from 'fs';
import * as JSZip from 'jszip';
import Automizer, { modify } from '../src/index';

/**
 * Regression test: hyperlinks created by a modification callback on an
 * APPENDED element (`slide.addElement()` + `modify.htmlToMultiText()`) must
 * survive the GenericShape hyperlink import.
 *
 * The callback runs during `GenericShape.prepare()` and mints its hyperlink
 * relationships on the TARGET slide's rels. `copyMultipleHyperlinks()` runs
 * afterwards and historically resolved every `a:hlinkClick r:id` against the
 * SOURCE slide's rels — the wrong id space for callback-created links. With
 * small source rels the ids simply resolved nowhere (and a naive
 * strip-unresolvable pass would delete the fresh links); with larger source
 * rels whatever relationship sat at those ids was cloned and the links
 * rewritten to it. The fix separates imported ids (present on the unmutated
 * source element → source id space) from callback-created ids (target id
 * space), so each is validated against the rels file it belongs to.
 */
test('htmlToMultiText links survive on an appended element', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<body><p>Intro</p>' +
    '<ol>' +
    '<li><a href="https://example.com/overview" target="_self"><strong>Bold link</strong></a></li>' +
    '</ol>' +
    '<p><a href="https://example.com/details" target="_self">Test link below</a></p></body>';

  const outputFile = `hyperlink-callback-links-append.test.pptx`;

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.addElement(
        'TextReplace.pptx',
        1,
        'setText',
        modify.htmlToMultiText(html),
      );
    })
    .write(outputFile);

  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${outputFile}`),
  );
  // With removeExistingSlides the root's own slide still serializes as
  // slide1; the added slide is slide2 — locate it by its content instead of
  // hardcoding, so the test survives numbering changes.
  let slideXml = '';
  let relsXml = '';
  for (const name of Object.keys(zip.files)) {
    const match = name.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (!match) continue;
    const xml = await zip.file(name)!.async('text');
    if (xml.includes('Test link below')) {
      slideXml = xml;
      relsXml = await zip
        .file(`ppt/slides/_rels/slide${match[1]}.xml.rels`)!
        .async('text');
      break;
    }
  }
  expect(slideXml).not.toBe('');

  // Both links must survive as hlinkClick elements whose r:id resolves to a
  // Relationship of Type .../hyperlink with the URL the callback set.
  const hlinkIds = [
    ...slideXml.matchAll(/<a:hlinkClick[^>]*r:id="([^"]+)"/g),
  ].map((m) => m[1]);
  expect(hlinkIds.length).toBeGreaterThanOrEqual(2);

  const relOf = (id: string) =>
    [...relsXml.matchAll(/<Relationship[^>]*>/g)]
      .map((m) => m[0])
      .find((r) => r.includes(`Id="${id}"`));

  for (const id of hlinkIds) {
    const rel = relOf(id);
    expect(rel).toBeDefined();
    expect(rel).toContain('relationships/hyperlink');
  }
  const targets = hlinkIds.map((id) => relOf(id));
  expect(targets.some((r) => r!.includes('example.com/overview'))).toBe(true);
  expect(targets.some((r) => r!.includes('example.com/details'))).toBe(true);
});
