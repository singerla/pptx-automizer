import Automizer, {ChartData, modify} from './index';
import {vd} from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithChartPoints.pptx`, 'charts');

const run = async () => {
  await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            {
              label: 'series 1',
              style: {
                color: {
                  type: 'schemeClr', value: 'accent1'
                }
              }
            },
            {label: 'series 2'},
            {label: 'series 3'},
          ],
          categories: [
            {
              label: 'cat 2-1',
              values: [50, 50, 20],
              styles: [
                {
                  color: {
                    type: 'srgbClr', value: '333333'
                  }
                }
              ]
            },
            {
              label: 'cat 2-2',
              values: [25, 10, 20],
              styles: [
                null,
                {
                  color: {
                    type: 'srgbClr', value: 'efefef'
                  }
                },
                {
                  color: {
                    type: 'srgbClr', value: 'eecc00'
                  }
                }
              ]
            },
            {label: 'cat 2-3', values: [15, 50, 20]},
            {label: 'cat 2-4', values: [26, 50, 20], styles: [
              null,
              null,
              {
                color: {
                  type: 'srgbClr', value: 'eeccff'
                }
              }
            ]}
          ],

        })
      ]);
    })
    .write(`modify-existing-chart-style.test.pptx`);

  return pres;
};

run().catch((error) => {
  console.error(error);
});
