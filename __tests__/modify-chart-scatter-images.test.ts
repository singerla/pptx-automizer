import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { ChartData, modify } from '../src/index';
import { expectXml } from './helpers/expect-xml';

test('create presentation, add and modify a scatter chart with embedded point images.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    // Marker images are referenced from the chart's own rels, not from the
    // slide; cleanup must keep them (ContentTracker collects chart images).
    cleanup: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartScatter.pptx`, 'charts');

  const dataScatter = <ChartData>(<unknown>{
    series: [{ label: 'series s1' }],
    categories: [
      { label: 'r1', values: [{ x: 10, y: 20 }] },
      { label: 'r2', values: [{ x: 21, y: 11 }] },
      { label: 'r3', values: [{ x: 22, y: 28 }] },
      { label: 'r4', values: [{ x: 13, y: 13 }] },
    ],
  });

  const outputFile = `modify-chart-scatter-images.test.pptx`;
  const result = await pres
    .addSlide('charts', 3, (slide) => {
      slide.modifyElement('ScatterPointImages', [
        modify.setChartScatter(dataScatter),
      ]);
    })
    .write(outputFile);

  expect(result.charts).toBe(2);

  // The marker image is related from ppt/charts/_rels only; it must survive
  // `cleanup: true` and the chart's rel must still point at the copy.
  const rels = await expectXml(outputFile, 'ppt/charts/_rels/chart1.xml.rels');
  rels.toHaveAttribute(
    'Relationship',
    'Target',
    '../media/image1.png-chart-1.png',
  );
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${outputFile}`),
  );
  expect(zip.file('ppt/media/image1.png-chart-1.png')).not.toBeNull();
});
