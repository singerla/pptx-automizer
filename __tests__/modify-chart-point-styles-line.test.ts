import * as fs from 'fs';
import JSZip from 'jszip';
import Automizer, { modify } from '../src/index';
import { ChartData, ChartValueStyle } from '../src/types/chart-types';
import { expectXml } from './helpers/expect-xml';
import { Element as XmlElement } from '@xmldom/xmldom';

/**
 * ROADMAP, <c:dPt> bug track — tier-0 guard.
 *
 * A data point that carries no applicable style must not produce a c:dPt at
 * all: fabricated points used to arrive with a default grey spPr whose
 * <a:ln><a:noFill/> erased the segments of line charts. No template ships a
 * line chart, so stage 1 generates one via pptxgenjs and stage 2 re-loads it
 * to run setChartData with per-point styles on it.
 */

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const CATEGORIES = ['alpha', 'beta', 'gamma', 'delta', 'epsilon'];
const SERIES_COUNT = 3;

const seriesValues = (s: number) =>
  CATEGORIES.map((_, c) => 2 + ((s * 5 + c * 3) % 11));

// Category 1 carries a real style; the rest are pseudo-styles that must not
// reach the XML layer: truthy keys without any applicable modification.
const styledCategories = (): Record<number, (ChartValueStyle | null)[]> => ({
  1: [{ color: { type: 'srgbClr', value: 'FF0000' } }],
  2: [{ marker: {} }],
  3: [{ color: <never>{ value: '00FF00' } }], // color without `type`
});

const dataPointsOf = (ser: XmlElement) =>
  Array.from(ser.getElementsByTagName('c:dPt')).map((dPt) => ({
    idx: dPt.getElementsByTagName('c:idx')[0].getAttribute('val'),
    color: dPt.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val'),
  }));

test('per-point styles on a line chart create no c:dPt beyond the explicitly styled points', async () => {
  const sourceFile = 'modify-chart-point-styles-line-source.pptx';

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
  const sourceArchive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${sourceFile}`),
  );
  const slideXml = await sourceArchive
    .file('ppt/slides/slide2.xml')
    .async('text');
  const chartName = slideXml.match(/name="(LineChart[^"]*)"/)?.[1];
  expect(chartName).toBeDefined();

  const outputFile = 'modify-chart-point-styles-line.test.pptx';
  const styles = styledCategories();

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
            styles: styles[c],
          })),
        }),
      ]);
    })
    .write(outputFile);

  // Locate the line chart part in the output (part numbering depends on the
  // charts the root template already carries).
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  const chartParts = Object.keys(archive.files).filter((path) =>
    /^ppt\/charts\/chart[0-9]+\.xml$/.test(path),
  );
  let linePart: string;
  for (const path of chartParts) {
    const xml = await archive.file(path).async('text');
    if (xml.includes('<c:lineChart>')) linePart = path;
  }
  expect(linePart).toBeDefined();

  const chart = await expectXml(outputFile, linePart);
  chart.toContainElement('c:v', 'series 1');

  const series = chart.elements('c:ser');
  expect(series.length).toBe(SERIES_COUNT);

  // Only the real style at category 1 materializes; the pseudo-styles at
  // categories 2 and 3 produce no c:dPt — not even an empty shell.
  expect(dataPointsOf(series[0])).toEqual([{ idx: '1', color: 'FF0000' }]);
  expect(dataPointsOf(series[1])).toEqual([]);
  expect(dataPointsOf(series[2])).toEqual([]);

  // The original defect: fabricated default spPr with <a:ln><a:noFill/>
  // erased line segments. No c:dPt may carry a noFill the caller never asked
  // for.
  for (const ser of series) {
    for (const dPt of Array.from(ser.getElementsByTagName('c:dPt'))) {
      expect(dPt.getElementsByTagName('a:noFill').length).toBe(0);
    }
  }
});
