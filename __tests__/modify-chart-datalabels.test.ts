import Automizer, { LabelPosition, modify } from '../src/index';
import { ChartSeriesDataLabelAttributes } from '../src/types/chart-types';

test('modify chart data label.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStackedLabels.pptx`, 'charts')
    .load(`ChartLinesVertical.pptx`, 'chartLines');

  const DataLabelAttributes: ChartSeriesDataLabelAttributes = {
    dLblPos: LabelPosition.Top,
    showLegendKey: false,
    showVal: false,
    showCatName: true,
    showSerName: false,
    showPercent: false,
    showBubbleSize: false,
    showLeaderLines: false,
  };

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setDataLabelAttributes(DataLabelAttributes),
      ]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setDataLabelAttributes({
          applyToSeries: 1,
          dLblPos: LabelPosition.InsideBase,
          showVal: true,
          showSerName: true,
          showPercent: true,
        }),
      ]);
    })
    .addSlide('chartLines', 1, (slide) => {
      slide.modifyElement('DotMatrix', [
        modify.setDataLabelAttributes({
          dLblPos: LabelPosition.Top,
          showLegendKey: true,
          showCatName: true,
          showSerName: true,
          solidFill: {
            type: 'srgbClr',
            value: '#FF00CC',
          },
        }),
      ]);
    })
    .write(`modify-chart-datalabels.test.pptx`);

  expect(result.charts).toBe(6);
});
