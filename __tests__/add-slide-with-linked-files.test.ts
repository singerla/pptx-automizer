import Automizer from '../src/automizer';
import { expectXml } from './helpers/expect-xml';
import JSZip from 'jszip';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * A slide with a picture and an OLE object that are linked to files
 * outside of the presentation (Insert > Picture > Link to File, Insert >
 * Object > Link): their relationships have TargetMode="External".
 */
const slideWithLinkedFiles = async (): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/threeSlidePres.pptx`),
  );

  const rels = 'ppt/slides/_rels/slide1.xml.rels';
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rId90" Type="${REL}/image" Target="file:///C:/linked/logo.png" TargetMode="External"/>` +
        `<Relationship Id="rId91" Type="${REL}/oleObject" Target="file:///C:/linked/data.xlsx" TargetMode="External"/>` +
        '</Relationships>',
    ),
  );

  const slide = 'ppt/slides/slide1.xml';
  zip.file(
    slide,
    (await zip.file(slide).async('string')).replace(
      '</p:spTree>',
      '<p:pic><p:nvPicPr><p:cNvPr id="90" name="Linked picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
        '<p:blipFill><a:blip r:link="rId90"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="91" name="Linked object"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
        '<p:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
        '<p:oleObj progId="Excel.Sheet.12" r:id="rId91"><p:link updateAutomatic="1"/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>' +
        '</p:spTree>',
    ),
  );

  return zip.generateAsync({ type: 'nodebuffer' });
};

test('append a slide with a linked picture and a linked OLE object', async () => {
  const automizer = new Automizer({ templateDir, outputDir });

  const result = await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(await slideWithLinkedFiles(), 'linked')
    .addSlide('linked', 1)
    .write(`add-slide-with-linked-files.test.pptx`);

  expect(result.slides).toBe(2);
  expect(result.images).toBe(0);

  // The links are kept as they are; there is nothing to copy.
  const rels = await expectXml(
    `add-slide-with-linked-files.test.pptx`,
    'ppt/slides/_rels/slide2.xml.rels',
  );
  const external = rels
    .elements('Relationship')
    .filter((rel) => rel.getAttribute('TargetMode') === 'External')
    .map((rel) => `${rel.getAttribute('Id')} ${rel.getAttribute('Target')}`);
  expect(external).toEqual([
    'rId90 file:///C:/linked/logo.png',
    'rId91 file:///C:/linked/data.xlsx',
  ]);

  const slide = await expectXml(
    `add-slide-with-linked-files.test.pptx`,
    'ppt/slides/slide2.xml',
  );
  slide.toHaveAttribute('a:blip', 'r:link', 'rId90');
  slide.toHaveAttribute('p:oleObj', 'r:id', 'rId91');
});
