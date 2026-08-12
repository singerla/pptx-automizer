import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import Automizer, { modify } from '../src/index';

test('delete internal hyperlink - using removeHyperlink helper', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    verbosity: 0
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `delete-hyperlink-internal.test.pptx`;

  const result = await pres
    .addSlide('link', 1, (slide) => {
      slide.modifyElement('LinkToSlide', [modify.removeHyperlink()]);
    })
    .write(outputFile);

  expect(result.slides).toBe(2);

  const zip = await JSZip.loadAsync(
    fs.readFileSync(path.join(`${__dirname}/pptx-output`, outputFile)),
  );
  const slideXml = await zip.file('ppt/slides/slide2.xml')!.async('text');
  const relsXml = await zip
    .file('ppt/slides/_rels/slide2.xml.rels')!
    .async('text');

  // The hyperlink of 'LinkToSlide' is gone, the ones of the other shapes
  // pointing to the same slide are untouched.
  const internalLinks = slideXml.match(/action="ppaction:\/\/hlinksldjump"/g);
  expect(internalLinks?.length).toBe(3);

  // Every r:id on the slide still resolves to a relationship, otherwise
  // PowerPoint asks to repair the file.
  const declaredRIds = Array.from(relsXml.matchAll(/Id="([^"]+)"/g)).map(
    (match) => match[1],
  );
  const usedRIds = Array.from(slideXml.matchAll(/r:(?:id|embed)="([^"]+)"/g))
    .map((match) => match[1]);

  expect(usedRIds.length).toBeGreaterThan(0);
  usedRIds.forEach((rId) => expect(declaredRIds).toContain(rId));
});
