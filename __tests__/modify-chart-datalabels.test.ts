import Automizer, {  LabelPosition, modify } from '../src/index';

test('modify chart data label.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStackedLabels.pptx`, 'charts');


  const DataLabelAttributes = {
    dLblPos: LabelPosition.Top,
    showLegendKey: false,
    showVal: false,
    showCatName: true,
    showSerName: false,
    showPercent: false,
    showBubbleSize: false,
    showLeaderLines: false
  }

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setDataLabelAttributes(DataLabelAttributes),
      ]);
    })
    .write(`modify-chart-datalabels.test.pptx`);

  expect(result.charts).toBe(2);
});
