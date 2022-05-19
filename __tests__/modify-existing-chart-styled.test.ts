import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';

test('create presentation, add slide with charts from template and modify existing chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            {
              label: 'series 1',
              // Style prop can be applied to series
              style: {
                color: {
                  type: 'schemeClr',
                  value: 'accent1',
                },
                // All labels of a series can be styled
                label: {
                  color: {
                    type: 'schemeClr',
                    value: 'accent2',
                  },
                  isBold: false,
                  size: 2200,
                },
              },
            },
            { label: 'series 2' },
            { label: 'series 3' },
          ],
          categories: [
            {
              label: 'cat 2-1',
              values: [50, 40, 20],
              // Style prop can be applied to single values,
              // array indices need to correspond (0: 50)
              styles: [
                {
                  color: {
                    type: 'srgbClr',
                    value: '333333',
                  },
                },
              ],
            },
            {
              label: 'cat 2-2',
              values: [25, 10, 20],
              // Style prop will be applied to second point in category ("10").
              styles: [
                null,
                {
                  color: {
                    type: 'srgbClr',
                    value: 'efefef',
                  },
                },
                {
                  color: {
                    type: 'srgbClr',
                    value: 'eecc00',
                  },
                },
              ],
            },
            { label: 'cat 2-3', values: [15, 50, 20] },
            {
              label: 'cat 2-4',
              values: [26, 50, 20],
              // Style prop will be applied to third point in category ("20").
              styles: [
                null,
                null,
                {
                  color: {
                    type: 'srgbClr',
                    value: 'eeccff',
                  },
                  // All single datapoint label can have a different style
                  label: {
                    color: {
                      type: 'schemeClr',
                      value: 'accent2',
                    },
                    isBold: false,
                    size: 2200,
                  },
                },
              ],
            },
          ],
        }),
      ]);
    })
    .write(`modify-existing-chart-styled.test.pptx`);

  expect(result.charts).toBe(3);
});
