import Automizer from '../src/automizer';
import { expectXml } from './helpers/expect-xml';

/**
 * Regression test for a chart's `chartUserShapes` drawing part losing its
 * own relationships when the chart is copied onto a new slide.
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
 * entry) on every `pres.write()`, so this test fails before the fix and
 * passes after it, without needing a custom assertion.
 */
test('chart userShapes drawing keeps its own image relationship when the chart is copied to a new slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartUserShapeWithImage.pptx`, 'chartUserShape');

  pres.addSlide('chartUserShape', 1);

  const result = await pres.write('chart-user-shape-rels.test.pptx');

  expect(result.charts).toBe(1);

  // Belt-and-suspenders: the copied drawing's own rels file must exist and
  // resolve the image relationship the invariants check would otherwise
  // report as dangling.
  const drawingRels = await expectXml(
    'chart-user-shape-rels.test.pptx',
    'ppt/drawings/_rels/drawing1.xml.rels',
  );
  drawingRels.toHaveAttribute(
    'Relationship',
    'Target',
    '../media/userShapeImage1.png-chart-1.png',
  );
});
