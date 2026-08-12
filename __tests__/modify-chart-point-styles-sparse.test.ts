import Automizer, { modify } from '../src/index';
import { ChartData, ChartValueStyle } from '../src/types/chart-types';
import { expectXml } from './helpers/expect-xml';
import { Element as XmlElement } from '@xmldom/xmldom';

/**
 * ROADMAP, Modification-contract track — tier-0 guards.
 *
 * c:dPt / c:dLbls are sparse collections: one element per *explicitly styled*
 * point, addressed by the <c:idx> payload, never by sibling position. Both
 * regressions of the track were invisible because no fixture exercised a
 * non-contiguous style set, and the only styled-chart test asserted nothing
 * about XML.
 */

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const categories = (styled: Record<number, ChartValueStyle>) =>
  Array.from({ length: 15 }, (_, c) => ({
    label: `cat ${c + 1}`,
    values: [c + 1, 16 - c, 10],
    styles: styled[c] ? [styled[c]] : undefined,
  }));

const dataPointsOf = (ser: XmlElement) =>
  Array.from(ser.getElementsByTagName('c:dPt')).map((dPt) => ({
    idx: dPt.getElementsByTagName('c:idx')[0].getAttribute('val'),
    color: dPt.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val'),
  }));

test('point styles on non-contiguous categories create exactly one c:dPt per styled point', async () => {
  const outputFile = 'modify-chart-point-styles-sparse.test.pptx';

  await new Automizer({ templateDir, outputDir })
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            { label: 'series 1' },
            { label: 'series 2' },
            { label: 'series 3' },
          ],
          // categories 0 and 3 of 15 styled — the sparse, non-contiguous
          // case that positional addressing silently collapsed into
          // duplicates of the first styled point
          categories: categories({
            0: { color: { type: 'srgbClr', value: 'FF0000' } },
            3: { color: { type: 'srgbClr', value: '00FF00' } },
          }),
        }),
      ]);
    })
    .write(outputFile);

  const chart = await expectXml(outputFile, 'ppt/charts/chart3.xml');
  chart.toContainElement('c:v', 'series 1');

  const series = chart.elements('c:ser');
  expect(series.length).toBe(3);

  expect(dataPointsOf(series[0])).toEqual([
    { idx: '0', color: 'FF0000' },
    { idx: '3', color: '00FF00' },
  ]);
  expect(dataPointsOf(series[1])).toEqual([]);
  expect(dataPointsOf(series[2])).toEqual([]);
});

test('a series style.label on a template without c:dLbls does not fabricate labels', async () => {
  const outputFile = 'modify-chart-no-fabricated-dlbls.test.pptx';

  await new Automizer({ templateDir, outputDir })
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            {
              label: 'series 1',
              style: {
                label: {
                  color: { type: 'schemeClr', value: 'accent2' },
                  size: 1400,
                },
              },
            },
            { label: 'series 2' },
            { label: 'series 3' },
          ],
          categories: categories({}),
        }),
      ]);
    })
    .write(outputFile);

  const chart = await expectXml(outputFile, 'ppt/charts/chart3.xml');
  chart.toContainElement('c:v', 'series 1');

  // "isRequired: false" guarantees modify-if-present: the template series
  // carry no c:dLbls, so none may appear — the chart-level c:dLbls of
  // <c:barChart> is untouched and stays the only one in the part.
  for (const ser of chart.elements('c:ser')) {
    expect(ser.getElementsByTagName('c:dLbls').length).toBe(0);
  }
  chart.toContainElementTimes('c:dLbls', 1);
});
