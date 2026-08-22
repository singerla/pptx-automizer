import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import Automizer from '../src/index';

/**
 * Regression test for a relationship-id collision: a hyperlink shape cloned
 * from a source slide carries its *source* r:id (here "rId3"), and the target
 * slide may already declare an unrelated relationship under that same id
 * (SlideWithImages slide 2 declares rId3 as an image). The element's r:id must
 * be rewritten to a freshly reserved id with its own hyperlink relationship —
 * skipping that rewrite would leave the hyperlink pointing at the image
 * relationship, an internally consistent mislink that PowerPoint rejects.
 */
test('addElement of a hyperlink shape onto a slide with a colliding rId', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images')
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `hyperlink-relid-collision.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  await pres
    .addSlide('images', 2, (slide) => {
      slide.addElement('link', 1, 'LinkToSlide');
    })
    .write(outputFile);

  const zip = await JSZip.loadAsync(fs.readFileSync(outputPath));
  const slideXml = await zip.file(`ppt/slides/slide2.xml`)!.async('text');
  const relsXml = await zip
    .file(`ppt/slides/_rels/slide2.xml.rels`)!
    .async('text');

  const rIdMatch = slideXml.match(
    /<p:cNvPr id="\d+" name="LinkToSlide"[^>]*>[\s\S]*?<a:hlinkClick r:id="([^"]+)"/,
  );
  expect(rIdMatch).not.toBeNull();
  const linkRid = rIdMatch![1];

  // The relationship the hyperlink resolves to must be a slide link — not the
  // image relationship that happens to sit at the shape's source rId.
  const relMatch = relsXml.match(
    new RegExp(`<Relationship\\s[^>]*Id="${linkRid}"[^>]*/?>`),
  );
  expect(relMatch).not.toBeNull();
  expect(relMatch![0]).toContain('relationships/slide"');
  expect(relMatch![0]).not.toContain('relationships/image');
});
