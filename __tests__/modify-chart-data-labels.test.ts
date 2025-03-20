import Automizer, { LabelPosition, modify } from '../src/index';

test('modify chart series data labels.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStackedLabels.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setDataLabelAttributes({
          dLblPos: LabelPosition.OutsideEnd,
        }),
      ]);
    })
    .write(`modify-chart-data-labels.test.pptx`);

  expect(result.charts).toBe(2);
});
