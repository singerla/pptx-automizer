import fs from 'fs';
import JSZip from 'jszip';
import Automizer from '../src/automizer';
import { expectXml } from './helpers/expect-xml';

/**
 * Regression tests for a chart's `chartUserShapes` drawing part losing its
 * own relationships when the chart is copied onto a new slide (#215).
 *
 * `ChartUserShapeWithImage.pptx` is `ChartUserShape.pptx` with a <cdr:pic>
 * added to the chart's userShapes drawing, referencing an embedded image via
 * its own ppt/drawings/_rels/drawing1.xml.rels — mirroring a think-cell
 * chart annotation that embeds a picture.
 *
 * Before the fix, `Chart.copyChartStyleFiles()` copied the drawing XML but
 * never copied its `_rels` file, so the copied drawing's `r:embed` pointed
 * nowhere in the output archive. `setup-pptx-invariants.ts` catches exactly
 * this class of defect (a relationship attribute with no matching `_rels`
 * entry) on every `pres.write()`; the assertions below pin the parts the
 * invariants would otherwise only report as missing.
 */

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const sourceImage = 'userShapeImage1.png';

/**
 * Asserts that the copied drawing `n` keeps a resolvable image: the r:embed
 * on its <cdr:pic> has an entry in the drawing's own rels, that entry points
 * at a media file that exists in the output zip, and the file's extension is
 * declared in [Content_Types].xml.
 */
const expectDrawingImage = async (
  outputFile: string,
  drawingNumber: number,
): Promise<string> => {
  const drawing = await expectXml(
    outputFile,
    `ppt/drawings/drawing${drawingNumber}.xml`,
  );
  const blip = drawing.elements('a:blip')[0];
  expect(blip).toBeDefined();
  const rId = blip.getAttribute('r:embed');
  expect(rId).toMatch(/^rId\d+/);

  const rels = await expectXml(
    outputFile,
    `ppt/drawings/_rels/drawing${drawingNumber}.xml.rels`,
  );
  rels.toHaveAttribute('Relationship', 'Id', rId);
  const relationship = rels
    .elements('Relationship')
    .find((element) => element.getAttribute('Id') === rId);
  const target = relationship.getAttribute('Target');
  expect(target).toMatch(/^\.\.\/media\//);

  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  const mediaPath = target.replace('../', 'ppt/');
  expect(zip.file(mediaPath)).not.toBeNull();

  const contentTypes = await expectXml(outputFile, '[Content_Types].xml');
  contentTypes.toHaveAttribute('Default', 'Extension', 'png');

  return target;
};

test('chart userShapes drawing keeps its own image relationship when the chart is copied to a new slide', async () => {
  const automizer = new Automizer({ templateDir, outputDir });
  const outputFile = 'chart-user-shape-rels.test.pptx';

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartUserShapeWithImage.pptx`, 'chartUserShape');

  pres.addSlide('chartUserShape', 1);

  const result = await pres.write(outputFile);

  expect(result.charts).toBe(1);

  const target = await expectDrawingImage(outputFile, 1);
  expect(target).toBe(`../media/${sourceImage}-chart-1.png`);
});

test('cloning the same chart twice gives each drawing its own rels and image', async () => {
  const automizer = new Automizer({ templateDir, outputDir });
  const outputFile = 'chart-user-shape-rels-twice.test.pptx';

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartUserShapeWithImage.pptx`, 'chartUserShape');

  pres.addSlide('chartUserShape', 1).addSlide('chartUserShape', 1);

  const result = await pres.write(outputFile);

  expect(result.charts).toBe(2);

  // The source rels must not be rewritten in place: the second copy has to
  // derive its image from the template's `../media/userShapeImage1.png`,
  // not from the already-renamed target of the first copy.
  const first = await expectDrawingImage(outputFile, 1);
  const second = await expectDrawingImage(outputFile, 2);
  expect(first).toBe(`../media/${sourceImage}-chart-1.png`);
  expect(second).toBe(`../media/${sourceImage}-chart-2.png`);
});

test('adding the chart as an element onto existing slides keeps the drawing image', async () => {
  const automizer = new Automizer({ templateDir, outputDir });
  const outputFile = 'chart-user-shape-rels-add-element.test.pptx';

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`ChartUserShapeWithImage.pptx`, 'chartUserShape');

  pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('chartUserShape', 1, 'BarsStackedWithUserShape');
    })
    .addSlide('empty', 1, (slide) => {
      slide.addElement('chartUserShape', 1, 'BarsStackedWithUserShape');
    });

  const result = await pres.write(outputFile);

  expect(result.charts).toBe(2);

  const first = await expectDrawingImage(outputFile, 1);
  const second = await expectDrawingImage(outputFile, 2);
  expect(first).toBe(`../media/${sourceImage}-chart-1.png`);
  expect(second).toBe(`../media/${sourceImage}-chart-2.png`);
});

test('the drawing image survives cleanup: true', async () => {
  const automizer = new Automizer({ templateDir, outputDir, cleanup: true });
  const outputFile = 'chart-user-shape-rels-cleanup.test.pptx';

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartUserShapeWithImage.pptx`, 'chartUserShape');

  pres.addSlide('chartUserShape', 1);

  const result = await pres.write(outputFile);

  expect(result.charts).toBe(1);

  const target = await expectDrawingImage(outputFile, 1);
  expect(target).toBe(`../media/${sourceImage}-chart-1.png`);
});
