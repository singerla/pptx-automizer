import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';
import { expectXml } from './helpers/expect-xml';
import { Element as XmlElement } from '@xmldom/xmldom';

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

  // The styles above are deliberately sparse: c:dPt/c:dLbl are one-element-
  // per-styled-point collections addressed by their <c:idx> payload, so each
  // series must end up with exactly the explicitly styled points — no
  // duplicated indices, no styles dropped (ROADMAP, Modification-contract
  // track, regression A).
  const chart = await expectXml(
    'modify-existing-chart-styled.test.pptx',
    'ppt/charts/chart3.xml',
  );
  // pin that chart3.xml is the modified ColumnChart
  chart.toContainElement('c:v', 'series 1');

  const series = chart.elements('c:ser');
  expect(series.length).toBe(3);

  const dataPointsOf = (ser: XmlElement) =>
    Array.from(ser.getElementsByTagName('c:dPt')).map((dPt) => ({
      idx: dPt.getElementsByTagName('c:idx')[0].getAttribute('val'),
      color: dPt.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val'),
    }));

  expect(dataPointsOf(series[0])).toEqual([{ idx: '0', color: '333333' }]);
  expect(dataPointsOf(series[1])).toEqual([{ idx: '1', color: 'EFEFEF' }]);
  expect(dataPointsOf(series[2])).toEqual([
    { idx: '1', color: 'EECC00' },
    { idx: '3', color: 'EECCFF' },
  ]);

  // series 1 carries a style.label, but the template series has no c:dLbls —
  // "modify if present" must not fabricate one (regression B).
  expect(series[0].getElementsByTagName('c:dLbls').length).toBe(0);

  // the single styled point label lands as a sparse c:dLbl at c:idx 3
  const pointLabels = Array.from(series[2].getElementsByTagName('c:dLbl'));
  expect(pointLabels.length).toBe(1);
  expect(
    pointLabels[0].getElementsByTagName('c:idx')[0].getAttribute('val'),
  ).toBe('3');
  const labelProps = pointLabels[0].getElementsByTagName('a:defRPr')[0];
  expect(labelProps.getAttribute('sz')).toBe('2200');
  expect(
    labelProps.getElementsByTagName('a:schemeClr')[0].getAttribute('val'),
  ).toBe('accent2');
});
