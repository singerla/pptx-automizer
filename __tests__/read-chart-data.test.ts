import Automizer, { LabelPosition, read } from '../src/index';

test('modify chart data label.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts');
  const data = [];

  await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [read.readWorkbookData(data)]);
    })
    .write(`read-chart-data.test.pptx`);

  expect(data.length).toBe(5);
  expect(data[0].length).toBe(4);
});
