import Automizer from '../src/automizer';
import { modify } from '../src';
import { expectXml } from './helpers/expect-xml';
import JSZip from 'jszip';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

/**
 * Some tools (e.g. .NET based ones such as Aspose.Slides) and a few of
 * PowerPoint's built-in themes write XML parts that start with a UTF-8
 * byte order mark (#164).
 */
const withByteOrderMarks = async (template: string): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/${template}`),
  );
  for (const name of Object.keys(zip.files).filter((name) =>
    /\.(xml|rels)$/.test(name),
  )) {
    zip.file(
      name,
      Buffer.concat([
        Buffer.from([0xef, 0xbb, 0xbf]),
        await zip.file(name).async('nodebuffer'),
      ]),
    );
  }
  return zip.generateAsync({ type: 'nodebuffer' });
};

test('load a root template and a template whose XML parts start with a BOM', async () => {
  const automizer = new Automizer({ templateDir, outputDir });

  const pres = automizer
    .loadRoot(await withByteOrderMarks(`RootTemplate.pptx`))
    .load(await withByteOrderMarks(`SlideWithShapes.pptx`), 'shapes');

  pres.addSlide('shapes', 1, (slide) => {
    slide.modifyElement('Titel 1', modify.setText('BOM'));
  });

  const result = await pres.write(`load-xml-with-bom.test.pptx`);
  expect(result.slides).toBe(2);

  (
    await expectXml(`load-xml-with-bom.test.pptx`, 'ppt/slides/slide2.xml')
  ).toContainElement('a:t', 'BOM');
});
