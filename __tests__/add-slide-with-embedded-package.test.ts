import Automizer from '../src/automizer';
import { XmlRelationshipHelper } from '../src/helper/xml-relationship-helper';
import { expectXml } from './helpers/expect-xml';
import JSZip from 'jszip';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const PACKAGE_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

/**
 * A slide with an embedded Excel workbook (Insert > Object), as PowerPoint
 * writes it: a p:oleObj related to the workbook by a `package`
 * relationship. The relationship is listed before the slide layout.
 */
const slideWithEmbeddedWorkbook = async (): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/threeSlidePres.pptx`),
  );
  zip.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx', 'workbook');

  const contentTypes = await zip.file('[Content_Types].xml').async('string');
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '<Default ',
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default ',
    ),
  );

  const rels = 'ppt/slides/_rels/slide1.xml.rels';
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '<Relationship ',
      `<Relationship Id="rId90" Type="${PACKAGE_REL}" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/><Relationship `,
    ),
  );

  const slide = 'ppt/slides/slide1.xml';
  zip.file(
    slide,
    (await zip.file(slide).async('string')).replace(
      '</p:spTree>',
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="90" name="Workbook"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
        '<p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
        '<p:oleObj progId="Excel.Sheet.12" r:id="rId90"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree>',
    ),
  );

  return zip.generateAsync({ type: 'nodebuffer' });
};

test('a package relationship only matches an embeddings prefix', () => {
  const embeddings = '../embeddings/Microsoft_Excel_Worksheet.xlsx';

  expect(
    XmlRelationshipHelper.targetMatchesRelationship(
      'package',
      '',
      embeddings,
      '../embeddings/',
    ),
  ).toBe(true);

  for (const prefix of [
    '../charts/chart',
    '../diagrams/data',
    '../slideLayouts/slideLayout',
  ]) {
    expect(
      XmlRelationshipHelper.targetMatchesRelationship(
        'package',
        prefix.split('/').pop(),
        embeddings,
        prefix,
      ),
    ).toBe(false);
  }
});

test('append a slide with an embedded Excel object', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    // copies the workbook, which is not handled by any shape type
    showIntegrityInfo: true,
  });

  const result = await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(await slideWithEmbeddedWorkbook(), 'workbook')
    .addSlide('workbook', 1)
    .write(`add-slide-with-embedded-package.test.pptx`);

  expect(result.slides).toBe(2);
  expect(result.charts).toBe(0);

  const rels = await expectXml(
    `add-slide-with-embedded-package.test.pptx`,
    'ppt/slides/_rels/slide2.xml.rels',
  );
  const targets = rels
    .elements('Relationship')
    .map(
      (rel) =>
        `${rel.getAttribute('Type').split('/').pop()} ${rel.getAttribute('Target')}`,
    );

  // The package relation still points at a workbook, not at a layout
  expect(targets).toContainEqual(
    expect.stringMatching(/^package \.\.\/embeddings\/.+\.xlsx$/),
  );
  expect(
    targets.filter((target) => target.startsWith('slideLayout ')),
  ).toHaveLength(1);
});
