import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import Automizer from '../src/index';

/**
 * Regression test for the "duplicate singleton relationship" corruption class
 * observed in field-generated decks: a multi-hyperlink shape (routed through
 * GenericShape → HyperlinkProcessor.copyMultipleHyperlinks) whose stale
 * a:hlinkClick r:id happens to match a *structural* relationship on the
 * source slide — classically rId1 = slideLayout or rId2 = notesSlide.
 *
 * copyMultipleHyperlinks used to clone whatever relationship sat at that id,
 * Type and all, giving the target slide a second slideLayout/notesSlide rel
 * (an OPC singleton violation: "can only have one instance of relationship
 * that targets part") with a source-numbered target. Only relationships of
 * Type hyperlink/slide may be copied; hyperlinks whose relationship cannot be
 * copied must be stripped, since a dangling r:id is itself a repair trigger.
 *
 * The fixture is built by patching SlideWithLink.pptx: the "ExternalLink"
 * shape's three a:hlinkClick r:ids are rewritten from rId2 (a genuine
 * hyperlink rel) to rId1 (the slideLayout rel).
 */

const outputDir = path.join(__dirname, 'pptx-output');
const patchedTemplate = 'stale-hyperlink-relid.generated.pptx';

beforeAll(async () => {
  const source = fs.readFileSync(
    path.join(__dirname, 'pptx-templates', 'SlideWithLink.pptx'),
  );
  const zip = await JSZip.loadAsync(source);
  const slidePath = 'ppt/slides/slide1.xml';
  const slideXml = await zip.file(slidePath)!.async('text');
  // Only the ExternalLink shape carries r:id="rId2" (three hlinkClicks).
  zip.file(slidePath, slideXml.split('r:id="rId2"').join('r:id="rId1"'));
  fs.writeFileSync(
    path.join(outputDir, patchedTemplate),
    await zip.generateAsync({ type: 'nodebuffer' }),
  );
});

test('stale hlinkClick r:id pointing at a structural rel is not cloned', async () => {
  const automizer = new Automizer({
    templateDir: outputDir,
    templateFallbackDir: path.join(__dirname, 'pptx-templates'),
    outputDir,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(patchedTemplate, 'stale');

  const outputFile = `hyperlink-stale-relid-generic.test.pptx`;
  const outputPath = path.join(outputDir, outputFile);

  await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('stale', 1, 'ExternalLink');
    })
    .write(outputFile);

  const zip = await JSZip.loadAsync(fs.readFileSync(outputPath));
  const slideXml = await zip.file(`ppt/slides/slide2.xml`)!.async('text');
  const relsXml = await zip
    .file(`ppt/slides/_rels/slide2.xml.rels`)!
    .async('text');

  // The shape itself must arrive on the slide, with its text intact.
  expect(slideXml).toContain('name="ExternalLink"');

  // Exactly one slideLayout relationship, and it is the slide's own — never
  // a copy created by the hyperlink import.
  const layoutRels = relsXml.match(
    /<Relationship\s[^>]*Type="[^"]*\/slideLayout"[^>]*>/g,
  );
  expect(layoutRels).toHaveLength(1);
  expect(layoutRels![0]).not.toContain('-created');

  // No notesSlide rel was conjured up either.
  expect(relsXml).not.toContain('/notesSlide"');

  // The uncopyable hyperlinks were stripped from the shape instead of being
  // left dangling or pointing at a structural relationship.
  const shapeMatch = slideXml.match(
    /<p:sp>(?:(?!<\/p:sp>)[\s\S])*name="ExternalLink"[\s\S]*?<\/p:sp>/,
  );
  expect(shapeMatch).not.toBeNull();
  expect(shapeMatch![0]).not.toContain('a:hlinkClick');
});
