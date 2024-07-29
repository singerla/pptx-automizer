import * as fs from 'fs/promises';
import Automizer from '../src/automizer';
import { ChartData, modify } from '../src';

test('create presentation from buffer/from directory and add basic slides', async () => {
  const automizer = new Automizer({
    outputDir: `${__dirname}/pptx-output`,
    templateDir: `${__dirname}/pptx-templates`,
  });
  const rootTemplate = await fs.readFile(
    `${__dirname}/pptx-templates/RootTemplate.pptx`,
  );
  const slideWithShapes = await fs.readFile(
    `${__dirname}/pptx-templates/SlideWithShapes.pptx`,
  );

  const url =
    'https://raw.githubusercontent.com/singerla/pptx-automizer/main/__tests__/pptx-templates/SlideWithShapes.pptx';

  const response = await fetch(url);
  const buffer = await response.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  const pres = automizer
    .loadRoot(rootTemplate)
    .load(bytes, 'shapesFromWeb')
    .load('SlideWithCharts.pptx', 'chartsFromDir')
    .load(slideWithShapes, 'shapes');

  pres.addSlide('chartsFromDir', 2, (slide) => {
    slide.modifyElement('ColumnChart', [
      modify.setChartData(<ChartData>{
        series: [{ label: 'series 1' }, { label: 'series 2' }],
        categories: [
          { label: 'cat 2-1', values: [50, 50] },
          { label: 'cat 2-2', values: [14, 50] },
          { label: 'cat 2-3', values: [15, 50] },
          { label: 'cat 2-4', values: [26, 50] },
        ],
      }),
    ]);
  });

  for (let i = 0; i <= 10; i++) {
    pres.addSlide('shapes', 1);
  }

  pres.addSlide('shapesFromWeb', 2);

  await pres.write(`create-presentation-from-buffer-mixed.test.pptx`);

  expect(pres).toBeInstanceOf(Automizer);
});
