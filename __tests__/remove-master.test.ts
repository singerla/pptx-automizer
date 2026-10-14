import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { AutomizerParams } from '../src';
import { expectXml } from './helpers/expect-xml';

/**
 * #85 — `pres.removeMaster(n)` removes a master of the root template with
 * its layouts, its theme (if unused elsewhere), its relationships and its
 * entries in presentation.xml and [Content_Types].xml.
 *
 * SlidesWithAdditionalMaster.pptx: master 1 ("Blue Design", layouts 1–12,
 * theme1 — also the presentation's theme), master 2 ("Orange Design",
 * layouts 13–20, theme2). Its slides 1 and 2 use layouts of master 2,
 * slide 3 uses layout 7 of master 1.
 */

const outputDir = `${__dirname}/pptx-output`;

const params = (extra: AutomizerParams = {}): AutomizerParams => ({
  templateDir: `${__dirname}/pptx-templates`,
  outputDir,
  verbosity: 0,
  ...extra,
});

async function partNames(outputFile: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(fs.readFileSync(`${outputDir}/${outputFile}`));
  return Object.keys(zip.files).filter((name) => !zip.files[name].dir);
}

const MIN_ID = 2147483648;

/** All p:sldMasterId / p:sldLayoutId ids of the output, in list order. */
async function masterAndLayoutIds(outputFile: string): Promise<number[]> {
  const presentation = await expectXml(outputFile, 'ppt/presentation.xml');
  const ids: number[] = [];
  for (const masterId of presentation.elements('p:sldMasterId')) {
    ids.push(Number(masterId.getAttribute('id')));
  }
  const parts = await partNames(outputFile);
  for (const part of parts) {
    if (!/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(part)) continue;
    const master = await expectXml(outputFile, part);
    for (const layoutId of master.elements('p:sldLayoutId')) {
      ids.push(Number(layoutId.getAttribute('id')));
    }
  }
  return ids;
}

test('removeMaster drops the master, its layouts, its theme and all entries', async () => {
  const output = 'remove-master.test.pptx';
  const pres = await new Automizer(params({ removeExistingSlides: true }))
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1)
    .removeMaster(2)
    .write(output);

  expect(pres.masters).toBe(1);

  const parts = await partNames(output);
  expect(parts).toContain('ppt/slideMasters/slideMaster1.xml');
  expect(parts).not.toContain('ppt/slideMasters/slideMaster2.xml');
  expect(parts).not.toContain('ppt/slideMasters/_rels/slideMaster2.xml.rels');
  for (let n = 13; n <= 20; n++) {
    expect(parts).not.toContain(`ppt/slideLayouts/slideLayout${n}.xml`);
    expect(parts).not.toContain(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`);
  }
  for (let n = 1; n <= 12; n++) {
    expect(parts).toContain(`ppt/slideLayouts/slideLayout${n}.xml`);
  }
  // theme2 belonged to master 2 only; theme1 is also the presentation's theme
  expect(parts).not.toContain('ppt/theme/theme2.xml');
  expect(parts).toContain('ppt/theme/theme1.xml');

  const presentation = await expectXml(output, 'ppt/presentation.xml');
  presentation.toContainElementTimes('p:sldMasterId', 1);

  const presentationRels = await expectXml(output, 'ppt/_rels/presentation.xml.rels');
  const relTargets = presentationRels
    .elements('Relationship')
    .map((rel) => rel.getAttribute('Target'));
  expect(relTargets).toContain('slideMasters/slideMaster1.xml');
  expect(relTargets).not.toContain('slideMasters/slideMaster2.xml');

  const contentTypes = await expectXml(output, '[Content_Types].xml');
  const overrides = contentTypes
    .elements('Override')
    .map((override) => override.getAttribute('PartName'));
  expect(overrides).not.toContain('/ppt/slideMasters/slideMaster2.xml');
  expect(overrides).not.toContain('/ppt/slideLayouts/slideLayout13.xml');
  expect(overrides).not.toContain('/ppt/theme/theme2.xml');
  expect(overrides).toContain('/ppt/slideMasters/slideMaster1.xml');

  // ids: all >= 2^31, unique, master id followed by its layout ids
  const ids = await masterAndLayoutIds(output);
  expect(ids).toHaveLength(1 + 12);
  expect(Math.min(...ids)).toBeGreaterThanOrEqual(MIN_ID);
  expect(new Set(ids).size).toBe(ids.length);
});

test('removing the first master keeps the ids of the remaining one valid', async () => {
  const output = 'remove-master-first.test.pptx';
  const pres = await new Automizer(params({ removeExistingSlides: true }))
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1, (slide) => {
      // the source slide uses layout 1 (master 1): switch it to master 2
      slide.useSlideLayout(13);
    })
    .removeMaster(1)
    .write(output);

  expect(pres.masters).toBe(1);

  const parts = await partNames(output);
  expect(parts).not.toContain('ppt/slideMasters/slideMaster1.xml');
  expect(parts).toContain('ppt/slideMasters/slideMaster2.xml');
  expect(parts).not.toContain('ppt/slideLayouts/slideLayout1.xml');
  expect(parts).toContain('ppt/slideLayouts/slideLayout13.xml');
  // theme1 is still referenced by ppt/_rels/presentation.xml.rels: kept
  expect(parts).toContain('ppt/theme/theme1.xml');

  // the root's three slides were dropped from the list (their parts stay
  // without cleanup); the added slide is slide4
  (
    await expectXml(output, 'ppt/slides/_rels/slide4.xml.rels')
  ).toHaveAttribute('Relationship', 'Target', '../slideLayouts/slideLayout13.xml');

  const ids = await masterAndLayoutIds(output);
  expect(ids).toHaveLength(1 + 8);
  expect(Math.min(...ids)).toBeGreaterThanOrEqual(MIN_ID);
  expect(new Set(ids).size).toBe(ids.length);
  // the remaining master's layout ids follow its own id
  expect(ids.slice(1).every((id) => id > ids[0])).toBe(true);
});

test('a master whose layout a slide still uses cannot be removed', async () => {
  // root slide 3 uses layout 7 of master 1
  await expect(
    new Automizer(params())
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .removeMaster(1)
      .write('remove-master-in-use.test.pptx'),
  ).rejects.toThrow('Cannot remove slide master 1: slide 3 uses slideLayout7');

  // an added slide keeps its source layout (layout 1 of master 1)
  await expect(
    new Automizer(params({ removeExistingSlides: true }))
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .load('SlideWithShapes.pptx')
      .addSlide('SlideWithShapes.pptx', 1)
      .removeMaster(1)
      .write('remove-master-in-use-added.test.pptx'),
  ).rejects.toThrow(/Cannot remove slide master 1: slide \d+ uses slideLayout1\b/);
});

test('removeMaster with cleanup: true', async () => {
  const output = 'remove-master-cleanup.test.pptx';
  const pres = await new Automizer(
    params({ removeExistingSlides: true, cleanup: true }),
  )
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .addSlide('SlideWithShapes.pptx', 1)
    .removeMaster(2)
    .write(output);

  expect(pres.masters).toBe(1);
  const parts = await partNames(output);
  expect(parts).not.toContain('ppt/slideMasters/slideMaster2.xml');
  expect(parts).not.toContain('ppt/slideLayouts/slideLayout20.xml');
  expect(parts).toContain('ppt/slideMasters/slideMaster1.xml');
  (await expectXml(output, 'ppt/presentation.xml')).toContainElementTimes(
    'p:sldMasterId',
    1,
  );
});

test('removing a master that does not exist rejects write()', async () => {
  await expect(
    new Automizer(params())
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .removeMaster(9)
      .write('remove-master-missing.test.pptx'),
  ).rejects.toThrow('ppt/slideMasters/slideMaster9.xml does not exist');
});
