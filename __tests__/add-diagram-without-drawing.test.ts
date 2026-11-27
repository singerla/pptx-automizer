import Automizer from '../src/automizer';
import { expectXml } from './helpers/expect-xml';
import JSZip from 'jszip';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

/**
 * The drawing part of a SmartArt diagram is a cached rendering, added by
 * PowerPoint 2010 and later. Older versions and other tools write the
 * diagram without it.
 */
const smartArtWithoutDrawing = async (): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/SmartArt.pptx`),
  );
  zip.remove('ppt/diagrams/drawing1.xml');

  const edit = async (part: string, edit: (xml: string) => string) =>
    zip.file(part, edit(await zip.file(part).async('string')));

  await edit('[Content_Types].xml', (xml) =>
    xml.replace(
      /<Override PartName="\/ppt\/diagrams\/drawing1\.xml"[^>]*>/,
      '',
    ),
  );
  await edit('ppt/slides/_rels/slide1.xml.rels', (xml) =>
    xml.replace(/<Relationship [^>]*diagramDrawing[^>]*>/, ''),
  );
  await edit('ppt/diagrams/data1.xml', (xml) =>
    xml.replace(/<dgm:extLst>.*<\/dgm:extLst>/, ''),
  );

  return zip.generateAsync({ type: 'nodebuffer' });
};

test('append a slide and a SmartArt diagram without a drawing part', async () => {
  const automizer = new Automizer({ templateDir, outputDir });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(await smartArtWithoutDrawing(), 'smartArt')
    .load(`EmptySlide.pptx`, 'empty');

  pres.addSlide('smartArt', 1);
  pres.addSlide('empty', 1, (slide) => {
    slide.addElement('smartArt', 1, '@SmartArt Hexagons');
  });

  const result = await pres.write(`add-diagram-without-drawing.test.pptx`);
  expect(result.slides).toBe(3);

  const output = `add-diagram-without-drawing.test.pptx`;
  for (const slide of [2, 3]) {
    const rels = await expectXml(
      output,
      `ppt/slides/_rels/slide${slide}.xml.rels`,
    );
    const types = rels
      .elements('Relationship')
      .map((rel) => rel.getAttribute('Type').split('/').pop());
    expect(types).toEqual(
      expect.arrayContaining([
        'diagramData',
        'diagramColors',
        'diagramLayout',
        'diagramQuickStyle',
      ]),
    );
    expect(types).not.toContain('diagramDrawing');
  }

  const contentTypes = await expectXml(output, '[Content_Types].xml');
  expect(contentTypes.raw()).toContain('/ppt/diagrams/data');
  expect(contentTypes.raw()).not.toContain('/ppt/diagrams/drawing');
});
