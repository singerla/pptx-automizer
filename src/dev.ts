import Automizer, { ChartData, modify } from './index';
import { vd } from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithChartsLabels.pptx`, 'charts');

const run = async () => {
  await pres
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            {
              label: 'series 1',
              // Style prop can be applied to series
              // style: {
              //   color: {
              //     type: 'schemeClr',
              //     value: 'accent2',
              //   },
              //   label: {
              //     color: {
              //       type: 'schemeClr',
              //       value: 'accent2',
              //     },
              //     isBold: false,
              //     size: 2200,
              //   },
              // },
            },
            {
              label: 'series 2',
              // style: {
              //   color: {
              //     type: 'schemeClr',
              //     value: 'accent1',
              //   },
              //   label: {
              //     color: {
              //       type: 'schemeClr',
              //       value: 'accent1',
              //     },
              //     isBold: true,
              //     size: 1200,
              //   },
              // },
            },
          ],
          categories: [
            {
              label: 'cat 2-1',
              values: [50, 40],
              styles: [
                null,
                {
                  // color: {
                  //   type: 'srgbClr',
                  //   value: '333333',
                  // },
                  label: {
                    color: {
                      type: 'schemeClr',
                      value: 'accent3',
                    },
                    isBold: true,
                    size: 5200,
                  },
                },
              ],
            },
            {
              label: 'cat 2-2',
              values: [25, 10],
              styles: [
                // null,
                // {
                //   color: {
                //     type: 'srgbClr',
                //     value: 'efefef',
                //   },
                //   label: {
                //     color: {
                //       type: 'schemeClr',
                //       value: 'accent1',
                //     },
                //     size: 3200,
                //   },
                // },
                // {
                //   color: {
                //     type: 'srgbClr',
                //     value: 'eecc00',
                //   },
                // },
              ],
            },
          ],
        }),
        // modify.dumpChart,
      ]);
    })
    .write(`modify-existing-chart-styled.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
