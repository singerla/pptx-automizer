import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { AutomizerParams, SlideLayoutNotFoundError } from '../src';
import { expectXml } from './helpers/expect-xml';

/**
 * #136 — `slide.useSlideLayout({ layout, master?, template? })` resolves
 * layouts by name against the output presentation, including the root
 * template's own layouts.
 *
 * Templates: RootTemplate.pptx has one master "Office" (layouts 1–11, e.g.
 * "Nur Titel" = 6, "Titel und Inhalt" = 2). SlidesWithAdditionalMaster.pptx
 * has the masters "Blue Design" (12 layouts) and "Orange Design" (8 layouts);
 * both also have a "Titel und Inhalt" and a "Nur Titel" layout.
 */

const outputDir = `${__dirname}/pptx-output`;

const automizerParams = (extra: AutomizerParams = {}): AutomizerParams => ({
  templateDir: `${__dirname}/pptx-templates`,
  outputDir,
  verbosity: 0,
  ...extra,
});

const layoutNumber = (layoutFile: string): number =>
  Number(layoutFile.match(/\d+/)[0]);

/** Layout file + its name + its master's theme name for an output slide. */
async function resolveSlideLayout(outputFile: string, slideNumber: number) {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  const read = (path: string) => zip.file(path).async('text');

  const slideRels = await read(`ppt/slides/_rels/slide${slideNumber}.xml.rels`);
  const layoutFile = slideRels.match(/slideLayouts\/(slideLayout\d+\.xml)/)[1];
  const layoutXml = await read(`ppt/slideLayouts/${layoutFile}`);
  const layoutName = layoutXml.match(/<p:cSld[^>]*name="([^"]*)"/)[1];

  const layoutRels = await read(`ppt/slideLayouts/_rels/${layoutFile}.rels`);
  const masterFile = layoutRels.match(/slideMasters\/(slideMaster\d+\.xml)/)[1];
  const masterRels = await read(`ppt/slideMasters/_rels/${masterFile}.rels`);
  const themeFile = masterRels.match(/theme\/(theme\d+\.xml)/)[1];
  const themeXml = await read(`ppt/theme/${themeFile}`);
  const masterName = themeXml.match(/<a:theme [^>]*name="([^"]*)"/)[1];

  return { layoutFile, layoutName, masterFile, masterName };
}

test('a root layout by name with template root imports nothing', async () => {
  const pres = await new Automizer(automizerParams())
    .loadRoot('RootTemplate.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1, (slide) => {
      slide.useSlideLayout({ layout: 'Nur Titel', template: 'root' });
    })
    .write('use-slide-layout-by-name-root.test.pptx');

  // the root master is the only one: nothing was auto-imported
  expect(pres.masters).toBe(1);
  (
    await expectXml(
      'use-slide-layout-by-name-root.test.pptx',
      'ppt/presentation.xml',
    )
  ).toContainElementTimes('p:sldMasterId', 1);

  // the root template brings its own slides; the added slide is the last
  (
    await expectXml(
      'use-slide-layout-by-name-root.test.pptx',
      `ppt/slides/_rels/slide${pres.slides}.xml.rels`,
    )
  ).toHaveAttribute('Relationship', 'Target', '../slideLayouts/slideLayout6.xml');

  const layout = await resolveSlideLayout(
    'use-slide-layout-by-name-root.test.pptx',
    pres.slides,
  );
  expect(layout).toMatchObject({
    layoutFile: 'slideLayout6.xml',
    layoutName: 'Nur Titel',
    masterName: 'Office',
  });
});

test('master + layout disambiguates a layout name shared by several masters', async () => {
  const pres = await new Automizer(automizerParams())
    .loadRoot('RootTemplate.pptx')
    .load('SlidesWithAdditionalMaster.pptx')
    .addMaster('SlidesWithAdditionalMaster.pptx', 1)
    .addMaster('SlidesWithAdditionalMaster.pptx', 2)
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      slide.useSlideLayout({
        master: 'Orange Design',
        layout: 'Titel und Inhalt',
      });
    })
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      slide.useSlideLayout({ master: 'Blue Design', layout: 'Titel und Inhalt' });
    })
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      slide.useSlideLayout({ master: 'Office', layout: 'Titel und Inhalt' });
    })
    .write('use-slide-layout-by-name-master.test.pptx');

  expect(pres.masters).toBe(3);

  // root: layouts 1–11, then the imported layouts of Blue Design (12–23)
  // and Orange Design (24–31), numbered in relationship order.
  const added = pres.slides - 2; // the three added slides are the last ones
  const orange = await resolveSlideLayout(
    'use-slide-layout-by-name-master.test.pptx',
    added,
  );
  expect(orange).toMatchObject({
    layoutName: 'Titel und Inhalt',
    masterName: 'Orange Design',
    masterFile: 'slideMaster3.xml',
  });
  expect(layoutNumber(orange.layoutFile)).toBeGreaterThanOrEqual(24);

  const blue = await resolveSlideLayout(
    'use-slide-layout-by-name-master.test.pptx',
    added + 1,
  );
  expect(blue).toMatchObject({
    layoutName: 'Titel und Inhalt',
    masterName: 'Blue Design',
    masterFile: 'slideMaster2.xml',
  });
  expect(layoutNumber(blue.layoutFile)).toBeGreaterThanOrEqual(12);
  expect(layoutNumber(blue.layoutFile)).toBeLessThanOrEqual(23);

  expect(
    await resolveSlideLayout(
      'use-slide-layout-by-name-master.test.pptx',
      added + 2,
    ),
  ).toMatchObject({
    layoutFile: 'slideLayout2.xml',
    layoutName: 'Titel und Inhalt',
    masterName: 'Office',
    masterFile: 'slideMaster1.xml',
  });
});

test('template restricts the lookup to layouts imported from that template', async () => {
  const pres = await new Automizer(automizerParams())
    .loadRoot('RootTemplate.pptx')
    .load('SlidesWithAdditionalMaster.pptx')
    .addMaster('SlidesWithAdditionalMaster.pptx', 2)
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // "Nur Titel" exists on the root master too; the template narrows it
      slide.useSlideLayout({
        layout: 'Nur Titel',
        template: 'SlidesWithAdditionalMaster.pptx',
      });
    })
    .write('use-slide-layout-by-name-template.test.pptx');

  // root: layouts 1–11, Orange Design: 12–19
  const layout = await resolveSlideLayout(
    'use-slide-layout-by-name-template.test.pptx',
    pres.slides,
  );
  expect(layout).toMatchObject({
    layoutName: 'Nur Titel',
    masterName: 'Orange Design',
    masterFile: 'slideMaster2.xml',
  });
  expect(layoutNumber(layout.layoutFile)).toBeGreaterThanOrEqual(12);
});

test('an ambiguous layout name rejects write() and lists the candidates', async () => {
  const pres = new Automizer(automizerParams())
    .loadRoot('RootTemplate.pptx')
    .load('SlidesWithAdditionalMaster.pptx')
    .addMaster('SlidesWithAdditionalMaster.pptx', 1)
    .addMaster('SlidesWithAdditionalMaster.pptx', 2)
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      slide.useSlideLayout({ layout: 'Titel und Inhalt' });
    });

  const write = pres.write('use-slide-layout-by-name-ambiguous.test.pptx');
  await expect(write).rejects.toBeInstanceOf(SlideLayoutNotFoundError);
  await expect(write).rejects.toThrow(/ambiguous/);
  await expect(write).rejects.toThrow('"Office / Titel und Inhalt" (slideLayout2, root)');
  await expect(write).rejects.toThrow(
    /"Blue Design \/ Titel und Inhalt" \(slideLayout\d+, SlidesWithAdditionalMaster\.pptx\)/,
  );
  await expect(write).rejects.toThrow(
    /"Orange Design \/ Titel und Inhalt" \(slideLayout\d+, SlidesWithAdditionalMaster\.pptx\)/,
  );
  await write.catch((e: SlideLayoutNotFoundError) => {
    expect(e.candidates.map((c) => c.masterName)).toEqual([
      'Office',
      'Blue Design',
      'Orange Design',
    ]);
    expect(e.candidates.map((c) => c.template)).toEqual([
      'root',
      'SlidesWithAdditionalMaster.pptx',
      'SlidesWithAdditionalMaster.pptx',
    ]);
  });
});

test('an unknown layout name rejects write() and lists what is available', async () => {
  const pres = new Automizer(automizerParams())
    .loadRoot('RootTemplate.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1, (slide) => {
      slide.useSlideLayout({ layout: 'No such layout' });
    });

  const write = pres.write('use-slide-layout-by-name-unknown.test.pptx');
  await expect(write).rejects.toBeInstanceOf(SlideLayoutNotFoundError);
  await expect(write).rejects.toThrow('layout "No such layout" not found');
  await expect(write).rejects.toThrow('"Office / Titelfolie" (slideLayout1, root)');
  await expect(write).rejects.toThrow('"Office / Nur Titel" (slideLayout6, root)');
});

test('continueOnError keeps the original layout on a failed selector', async () => {
  const pres = await new Automizer(automizerParams({ continueOnError: true }))
    .loadRoot('RootTemplate.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1, (slide) => {
      slide.useSlideLayout({ layout: 'No such layout' });
    })
    .write('use-slide-layout-by-name-lenient.test.pptx');

  const layout = await resolveSlideLayout(
    'use-slide-layout-by-name-lenient.test.pptx',
    pres.slides,
  );
  // the source slide's own layout number (slideLayout1 in SlideWithShapes.pptx)
  expect(layout.layoutFile).toBe('slideLayout1.xml');
});

test('a plain layout name falls back to the root template layouts', async () => {
  // The root's "Titelfolie blue text" exists on no other template, so the
  // auto-imported master of the source slide cannot provide it.
  const pres = await new Automizer(automizerParams())
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1, (slide) => {
      slide.useSlideLayout('Titelfolie blue text');
    })
    .write('use-slide-layout-by-name-string-fallback.test.pptx');

  // the string form still auto-imports the source slide's master
  expect(pres.masters).toBe(3);

  // this root template brings its own slides; the added slide is the last
  expect(
    await resolveSlideLayout(
      'use-slide-layout-by-name-string-fallback.test.pptx',
      pres.slides,
    ),
  ).toMatchObject({
    layoutFile: 'slideLayout1.xml',
    layoutName: 'Titelfolie blue text',
    masterName: 'Blue Design',
  });
});
