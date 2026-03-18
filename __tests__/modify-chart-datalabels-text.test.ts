import Automizer, { ChartData, modify } from '../src/index';

test('modify chart data and data labels style and text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStackedLabels.pptx`, 'charts');

  const dataSmaller: ChartData = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      // a null value will be converted to ""
      {
        label: 'item test r1',
        values: [11, 22],
        styles: [
          {
            color: {
              type: 'srgbClr',
              value: '#FF0000',
              alpha: 0.5
            },
            label: {
              suffix: {
                text: '**',
                color: {
                  type: 'srgbClr',
                  value: '#FF0000',
                },
              },
            },
          },
          {
            label: {
              suffix: {
                text: 'tmptmp',
              },
            },
          },
        ],
      },
      {
        label: 'item test r2',
        values: [12, 18],
        styles: [
          {
            label: {
              suffix: {
                text: '*',
                color: {
                  type: 'srgbClr',
                  value: '#FF0000',
                },
              },
            },
          },
          {
            label: {
              suffix: {
                text: 'tmp2',
              },
            },
          },
        ],
      },
    ],
  };

  const result = await pres
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('BarsStackedFmtLabels', [modify.setChartData(dataSmaller)]);
    })
    .write(`modify-chart-datalabels-text.test.pptx`);

  // expect(result.charts).toBe(4);
});
