import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { AutomizerParams, ModifyTextHelper } from '../src';
import { expectXml } from './helpers/expect-xml';

/**
 * #85 — `pres.modifyMaster(n, cb)` and `pres.modifyLayout(n, cb)` modify a
 * master/layout that is already part of the root template in place: no
 * copy is made and the part count stays the same.
 *
 * SlidesWithAdditionalMaster.pptx: master 1 ("Blue Design", layouts 1–12,
 * shapes "MasterRectangle" and "Rechteck 6"), master 2 ("Orange Design",
 * layouts 13–20, shape "Rechteck 6"). Layout 1 has the title placeholder
 * "Titel 1".
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
  return Object.keys(zip.files);
}

test('modifyMaster edits, adds and removes shapes on a root master in place', async () => {
  const output = 'modify-master-in-place.test.pptx';
  const pres = await new Automizer(params())
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .modifyMaster(1, (master) => {
      master.modifyElement(
        'MasterRectangle',
        ModifyTextHelper.setText('modified in place'),
      );
    })
    // a second call for the same master merges into the same queue
    .modifyMaster(1, (master) => {
      master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
    })
    .modifyMaster(2, (master) => {
      master.removeElement('Rechteck 6');
    })
    .addSlide('SlideWithShapes.pptx', 1)
    .write(output);

  // no master was copied
  expect(pres.masters).toBe(2);
  (await expectXml(output, 'ppt/presentation.xml')).toContainElementTimes(
    'p:sldMasterId',
    2,
  );
  const parts = await partNames(output);
  expect(parts.filter((part) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(part))).toHaveLength(2);
  expect(parts.filter((part) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(part))).toHaveLength(20);

  const master1 = await expectXml(output, 'ppt/slideMasters/slideMaster1.xml');
  master1.toContainElement('a:t', 'modified in place');
  master1.toHaveAttribute('p:cNvPr', 'name', 'MasterRectangle');
  master1.toHaveAttribute('p:cNvPr', 'name', 'Cloud 1');

  const master2 = await expectXml(output, 'ppt/slideMasters/slideMaster2.xml');
  expect(
    master2
      .elements('p:cNvPr')
      .map((element) => element.getAttribute('name')),
  ).not.toContain('Rechteck 6');
  // the other master kept its rectangle
  master1.toHaveAttribute('p:cNvPr', 'name', 'Rechteck 6');
});

test('modifyLayout edits a root layout in place', async () => {
  const output = 'modify-layout-in-place.test.pptx';
  const pres = await new Automizer(params())
    .loadRoot('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .modifyLayout(1, (layout) => {
      layout.modifyElement('Titel 1', ModifyTextHelper.setText('layout title'));
      layout.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
    })
    .modifyLayout(13, (layout) => {
      layout.removeElement('Untertitel 2');
    })
    .addSlide('SlideWithShapes.pptx', 1)
    .write(output);

  expect(pres.masters).toBe(2);
  const parts = await partNames(output);
  expect(parts.filter((part) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(part))).toHaveLength(20);

  const layout1 = await expectXml(output, 'ppt/slideLayouts/slideLayout1.xml');
  layout1.toContainElement('a:t', 'layout title');
  layout1.toHaveAttribute('p:cNvPr', 'name', 'Cloud 1');

  const layout13 = await expectXml(output, 'ppt/slideLayouts/slideLayout13.xml');
  expect(
    layout13
      .elements('p:cNvPr')
      .map((element) => element.getAttribute('name')),
  ).not.toContain('Untertitel 2');
  layout13.toHaveAttribute('p:cNvPr', 'name', 'Titel 1');
});

test('modifying a master or layout that does not exist rejects write()', async () => {
  await expect(
    new Automizer(params())
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .modifyMaster(3, (master) => {
        master.modifyElement('MasterRectangle', ModifyTextHelper.setText('x'));
      })
      .write('modify-master-in-place-missing.test.pptx'),
  ).rejects.toThrow('ppt/slideMasters/slideMaster3.xml does not exist');

  await expect(
    new Automizer(params())
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .modifyLayout(21, (layout) => {
        layout.modifyElement('Titel 1', ModifyTextHelper.setText('x'));
      })
      .write('modify-layout-in-place-missing.test.pptx'),
  ).rejects.toThrow('ppt/slideLayouts/slideLayout21.xml does not exist');
});

test('an unknown shape on a root master rejects write() like on a slide', async () => {
  await expect(
    new Automizer(params())
      .loadRoot('SlidesWithAdditionalMaster.pptx')
      .modifyMaster(1, (master) => {
        master.modifyElement('NoSuchShape', ModifyTextHelper.setText('x'));
      })
      .write('modify-master-in-place-unknown-shape.test.pptx'),
  ).rejects.toThrow("Can't find element");
});
