import * as fs from 'fs';
import JSZip from 'jszip';
import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';
import { ILogger } from '../src/helper/logger';
import { expectXml } from './helpers/expect-xml';
import { Element as XmlElement } from '@xmldom/xmldom';

/**
 * Tier-0 guard for `ModifyXmlHelper.createElement('c:pt')`.
 *
 * A chart whose caches hold no `<c:pt>` at all (PowerPoint writes such
 * "legend-only" charts; a template may also ship a stripped cache) could
 * never receive a point: `assertElement` clones a missing point from its
 * last sibling, and with no sibling `createElement` did not know the tag.
 * Every `point()` call then logged "Could not assert required tag: c:pt"
 * and the cache stayed empty, while the embedded workbook was written
 * correctly — PowerPoint renders from the cache, so the series showed
 * nothing until "Edit Data" was opened.
 *
 * Stage 1 builds the fixture by stripping all points from chart1.xml of
 * SlideWithCharts.pptx; stage 2 runs setChartData against it.
 */

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const CATEGORIES = ['alpha', 'beta', 'gamma'];
const SERIES = ['series 1', 'series 2'];
const VALUES = [
  [11, 21],
  [12, 22],
  [13, 23],
];

class CollectingLogger implements ILogger {
  warnings: string[] = [];
  error(): void {}
  warn(message: unknown): void {
    this.warnings.push(String(message));
  }
  info(): void {}
  debug(): void {}
}

const pointsOf = (wrapper: XmlElement) =>
  Array.from(wrapper.getElementsByTagName('c:pt')).map((pt) => ({
    idx: pt.getAttribute('idx'),
    value: pt.getElementsByTagName('c:v')[0]?.textContent,
  }));

test('setChartData fills a chart cache that holds no c:pt at all', async () => {
  const sourceFile = 'modify-chart-empty-cache-source.pptx';
  const sourceChartPart = 'ppt/charts/chart1.xml';

  // Stage 1: the fixture. Remove every point from c:cat and c:val and reset
  // their counts, as a legend-only chart saved by PowerPoint looks. The
  // series-name cache under c:tx keeps its point: PowerPoint needs it for
  // the legend, and `seriesLabel` addresses the first c:v of the series.
  const source = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/SlideWithCharts.pptx`),
  );
  const chartXml = await source.file(sourceChartPart).async('text');
  expect(chartXml).toMatch(/<c:pt idx="/);
  const stripped = chartXml.replace(
    /<c:(cat|val)>.*?<\/c:\1>/g,
    (wrapper) =>
      wrapper
        .replace(/<c:pt idx="\d+">.*?<\/c:pt>/g, '')
        .replace(/<c:ptCount val="\d+"\/>/g, '<c:ptCount val="0"/>'),
  );
  expect(stripped.match(/<c:pt /g).length).toBe(
    chartXml.match(/<c:tx>/g).length,
  );
  source.file(sourceChartPart, stripped);
  fs.writeFileSync(
    `${outputDir}/${sourceFile}`,
    await source.generateAsync({ type: 'nodebuffer' }),
  );

  // Stage 2: fill the stripped chart.
  const logger = new CollectingLogger();
  const outputFile = 'modify-chart-empty-cache.test.pptx';

  await new Automizer({
    templateDir: outputDir,
    templateFallbackDir: templateDir,
    outputDir,
    logger,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(sourceFile, 'charts')
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('StackedBars', [
        modify.setChartData(<ChartData>{
          series: SERIES.map((label) => ({ label })),
          categories: CATEGORIES.map((label, c) => ({
            label,
            values: VALUES[c],
          })),
        }),
      ]);
    })
    .write(outputFile);

  const pointWarnings = logger.warnings.filter((message) =>
    message.includes('Could not assert required tag: c:pt'),
  );
  expect(pointWarnings).toEqual([]);

  // The modified chart is a fresh part; locate it by the data it carries.
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  const chartParts = Object.keys(archive.files).filter((path) =>
    /^ppt\/charts\/chart[0-9]+\.xml$/.test(path),
  );
  let chartPart: string;
  for (const path of chartParts) {
    const xml = await archive.file(path).async('text');
    if (xml.includes('alpha')) chartPart = path;
  }
  expect(chartPart).toBeDefined();

  const chart = await expectXml(outputFile, chartPart);
  const series = chart.elements('c:ser');
  expect(series.length).toBe(SERIES.length);

  series.forEach((ser, s) => {
    const cat = ser.getElementsByTagName('c:cat')[0];
    const val = ser.getElementsByTagName('c:val')[0];

    expect(pointsOf(cat)).toEqual(
      CATEGORIES.map((label, c) => ({ idx: String(c), value: label })),
    );
    expect(pointsOf(val)).toEqual(
      VALUES.map((row, c) => ({ idx: String(c), value: String(row[s]) })),
    );

    for (const wrapper of [cat, val]) {
      const ptCount = wrapper.getElementsByTagName('c:ptCount')[0];
      expect(ptCount.getAttribute('val')).toBe(String(CATEGORIES.length));
      const cache = ptCount.parentNode as XmlElement;
      // Schema order inside the cache: formatCode?, ptCount?, pt*, extLst?
      const children = Array.from(cache.childNodes)
        .filter((node) => node.nodeType === 1)
        .map((node) => node.nodeName);
      const firstPt = children.indexOf('c:pt');
      const lastPt = children.lastIndexOf('c:pt');
      expect(firstPt).toBeGreaterThan(children.indexOf('c:ptCount'));
      expect(children.slice(firstPt, lastPt + 1)).toEqual(
        Array(CATEGORIES.length).fill('c:pt'),
      );
      const extLst = children.indexOf('c:extLst');
      if (extLst !== -1) expect(extLst).toBeGreaterThan(lastPt);
    }
  });
});

test('setExtendedChartData fills a chartex cx:lvl that holds no cx:pt at all', async () => {
  const sourceFile = 'modify-chart-empty-cache-ext-source.pptx';
  const sourceChartPart = 'ppt/charts/chartEx1.xml';

  // Stage 1: strip every cx:pt from both dimensions of the waterfall.
  const source = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/ChartWaterfall.pptx`),
  );
  const chartXml = await source.file(sourceChartPart).async('text');
  expect(chartXml).toMatch(/<cx:pt idx="/);
  const stripped = chartXml
    .replace(/<cx:pt idx="\d+">.*?<\/cx:pt>/g, '')
    .replace(/ptCount="\d+"/g, 'ptCount="0"');
  expect(stripped).not.toMatch(/<cx:pt /);
  source.file(sourceChartPart, stripped);
  fs.writeFileSync(
    `${outputDir}/${sourceFile}`,
    await source.generateAsync({ type: 'nodebuffer' }),
  );

  // Stage 2
  const logger = new CollectingLogger();
  const outputFile = 'modify-chart-empty-cache-ext.test.pptx';

  await new Automizer({
    templateDir: outputDir,
    templateFallbackDir: templateDir,
    outputDir,
    logger,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(sourceFile, 'charts')
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Waterfall 1', [
        modify.setExtendedChartData({
          series: [{ label: 'series 1' }],
          categories: CATEGORIES.map((label, c) => ({
            label,
            values: [VALUES[c][0]],
          })),
        }),
      ]);
    })
    .write(outputFile);

  const pointWarnings = logger.warnings.filter((message) =>
    message.includes('Could not assert required tag: cx:pt'),
  );
  expect(pointWarnings).toEqual([]);

  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  // The modified chart is a fresh part; locate it by the data it carries.
  let chartPart: string;
  for (const path of Object.keys(archive.files)) {
    if (!/^ppt\/charts\/chartEx[0-9]+\.xml$/.test(path)) continue;
    const xml = await archive.file(path).async('text');
    if (xml.includes('alpha')) chartPart = path;
  }
  expect(chartPart).toBeDefined();

  const chart = await expectXml(outputFile, chartPart);
  const extPointsOf = (dim: XmlElement) =>
    Array.from(dim.getElementsByTagName('cx:pt')).map((pt) => ({
      idx: pt.getAttribute('idx'),
      value: pt.textContent,
    }));

  const strDim = chart.elements('cx:strDim')[0];
  const numDim = chart.elements('cx:numDim')[0];
  expect(extPointsOf(strDim)).toEqual(
    CATEGORIES.map((label, c) => ({ idx: String(c), value: label })),
  );
  expect(extPointsOf(numDim)).toEqual(
    VALUES.map((row, c) => ({ idx: String(c), value: String(row[0]) })),
  );
  for (const dim of [strDim, numDim]) {
    expect(
      dim.getElementsByTagName('cx:lvl')[0].getAttribute('ptCount'),
    ).toBe(String(CATEGORIES.length));
  }
});
