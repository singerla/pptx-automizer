/**
 * Golden deck `chart-line-points` (Tier 3): sparse per-point styles applied
 * to a multi-series line chart.
 *
 * Guards the fabricated-c:dPt regression (ROADMAP, <c:dPt> bug track): a
 * point style that yields no applicable modification must not create a
 * c:dPt, and a created one must carry no default spPr — the former grey
 * <a:solidFill> + <a:ln><a:noFill/> blob removed the line segments, and the
 * chart rendered as floating data labels. Exactly the failure mode a pixel
 * diff catches and an XML diff understates.
 *
 * No template ships a line chart, so stage 1 generates one via pptxgenjs
 * and stage 2 re-loads that deck as a template to run setChartData on it.
 */
import * as fs from 'fs';
import JSZip from 'jszip';
import Automizer, { ChartData, modify } from '../../src/index';
import { ChartValueStyle } from '../../src/types/chart-types';
import { expectDeckMatchesBaselines } from './helpers/golden-deck';

const templateDir = `${__dirname}/../pptx-templates`;
const outputDir = `${__dirname}/../pptx-output`;

const CATEGORIES = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta'];
const SERIES_COUNT = 5;

const seriesValues = (s: number) =>
  CATEGORIES.map((_, c) => 2 + ((s * 5 + c * 3) % 11));

// Per category: one real fill, one real border, and pseudo-styles (truthy
// keys without applicable content) that must render as a plain line chart.
const styles = (c: number): (ChartValueStyle | null)[] | undefined => {
  switch (c) {
    case 1:
      return [{ color: { type: 'srgbClr', value: 'FF0000' } }];
    case 2:
      return [
        null,
        { border: { color: { type: 'srgbClr', value: '00AA00' }, weight: 40000 } },
      ];
    case 3:
      return Array.from({ length: SERIES_COUNT }, () => ({ marker: {} }));
    case 4:
      return [null, null, { color: <never>{ value: '0000FF' } }];
    default:
      return undefined;
  }
};

test('golden deck: chart-line-points', async () => {
  const sourceFile = 'visual-chart-line-points-source.pptx';

  await new Automizer({ templateDir, outputDir })
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((pSlide, pptxGenJs) => {
        pSlide.addChart(
          pptxGenJs.ChartType.line,
          Array.from({ length: SERIES_COUNT }, (_, s) => ({
            name: `series ${s + 1}`,
            labels: CATEGORIES,
            values: seriesValues(s),
          })),
          { x: 0.5, y: 0.5, w: 9, h: 6 },
        );
      }, 'LineChart');
    })
    .write(sourceFile);

  // slide.generate appends a uuid to the shape name — read it back to
  // address the chart in stage 2.
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${sourceFile}`),
  );
  const slideXml = await archive.file('ppt/slides/slide2.xml').async('text');
  const chartName = slideXml.match(/name="(LineChart[^"]*)"/)?.[1];
  expect(chartName).toBeDefined();

  const outputFile = 'visual-chart-line-points.pptx';
  await new Automizer({
    templateDir: outputDir,
    templateFallbackDir: templateDir,
    outputDir,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(sourceFile, 'line')
    .addSlide('line', 2, (slide) => {
      slide.modifyElement(chartName, [
        modify.setChartData(<ChartData>{
          series: Array.from({ length: SERIES_COUNT }, (_, s) => ({
            label: `series ${s + 1}`,
          })),
          categories: CATEGORIES.map((label, c) => ({
            label,
            values: Array.from({ length: SERIES_COUNT }, (_, s) =>
              seriesValues(s)[c],
            ),
            styles: styles(c),
          })),
        }),
      ]);
    })
    .write(outputFile);

  await expectDeckMatchesBaselines('chart-line-points', outputFile, 2);
});
