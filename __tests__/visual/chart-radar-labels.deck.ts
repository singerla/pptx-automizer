/**
 * Golden deck `chart-radar-labels` (Tier 3): styled series labels applied to
 * a multi-series radar chart whose template carries no data labels.
 *
 * Guards the fabricated-c:dLbls regression (ROADMAP, Modification-contract
 * track, B): `style.label` means "style labels where the template has them",
 * so the render must show a plain radar chart — a deck sprouting one default
 * label per series is exactly the customer-visible defect an XML diff
 * understates.
 *
 * No template ships a radar chart, so stage 1 generates one via pptxgenjs
 * and stage 2 re-loads that deck as a template to run setChartData on it.
 */
import * as fs from 'fs';
import JSZip from 'jszip';
import Automizer, { ChartData, modify } from '../../src/index';
import { expectDeckMatchesBaselines } from './helpers/golden-deck';

const templateDir = `${__dirname}/../pptx-templates`;
const outputDir = `${__dirname}/../pptx-output`;

const CATEGORIES = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta'];
const SERIES_COUNT = 5;

const seriesValues = (s: number) =>
  CATEGORIES.map((_, c) => 2 + ((s * 5 + c * 3) % 11));

test('golden deck: chart-radar-labels', async () => {
  const sourceFile = 'visual-chart-radar-source.pptx';

  await new Automizer({ templateDir, outputDir })
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((pSlide, pptxGenJs) => {
        pSlide.addChart(
          pptxGenJs.ChartType.radar,
          Array.from({ length: SERIES_COUNT }, (_, s) => ({
            name: `series ${s + 1}`,
            labels: CATEGORIES,
            values: seriesValues(s),
          })),
          { x: 0.5, y: 0.5, w: 9, h: 6, radarStyle: 'standard' },
        );
      }, 'RadarChart');
    })
    .write(sourceFile);

  // slide.generate appends a uuid to the shape name — read it back to
  // address the chart in stage 2.
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${sourceFile}`),
  );
  const slideXml = await archive.file('ppt/slides/slide2.xml').async('text');
  const chartName = slideXml.match(/name="(RadarChart[^"]*)"/)?.[1];
  expect(chartName).toBeDefined();

  const outputFile = 'visual-chart-radar-labels.pptx';
  await new Automizer({
    templateDir: outputDir,
    templateFallbackDir: templateDir,
    outputDir,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(sourceFile, 'radar')
    .addSlide('radar', 2, (slide) => {
      slide.modifyElement(chartName, [
        modify.setChartData(<ChartData>{
          series: Array.from({ length: SERIES_COUNT }, (_, s) => ({
            label: `series ${s + 1}`,
            style: {
              label: {
                color: { type: 'schemeClr' as const, value: 'accent2' },
                size: 1400,
              },
            },
          })),
          categories: CATEGORIES.map((label, c) => ({
            label,
            values: Array.from({ length: SERIES_COUNT }, (_, s) =>
              seriesValues(s)[c],
            ),
          })),
        }),
      ]);
    })
    .write(outputFile);

  await expectDeckMatchesBaselines('chart-radar-labels', outputFile, 2);
});
