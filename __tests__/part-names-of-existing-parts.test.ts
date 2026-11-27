import Automizer from '../src/automizer';
import JSZip from 'jszip';
import fs from 'fs';

/**
 * New charts, images, diagrams and OLE objects are named after a counter
 * (`chart<n>.xml`, `media<n>.png`, `data<n>.xml`, `oleObject<n>.bin`). The
 * counter has to start above every number already used by a part of the
 * root template - a part copied to a name that is taken overwrites a part
 * another slide still refers to.
 */

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const PART_FOLDERS =
  /^ppt\/(charts|media|diagrams|embeddings|drawings|theme)\/[^/]+$/;

/**
 * Contents of the chart, media, diagram and embedded parts of a .pptx.
 */
const readParts = async (
  file: Buffer | string,
): Promise<Record<string, string>> => {
  const zip = await JSZip.loadAsync(
    typeof file === 'string' ? fs.readFileSync(`${outputDir}/${file}`) : file,
  );
  const parts: Record<string, string> = {};
  for (const name of Object.keys(zip.files).filter((name) =>
    PART_FOLDERS.test(name),
  )) {
    parts[name] = await zip.file(name).async('base64');
  }
  return parts;
};

/**
 * Every part of the root template is still there, unchanged.
 */
const expectPartsKept = (
  before: Record<string, string>,
  after: Record<string, string>,
) => {
  const changed = Object.keys(before).filter(
    (name) => after[name] !== before[name],
  );
  expect(changed).toEqual([]);
  expect(Object.keys(after).length).toBeGreaterThan(Object.keys(before).length);
};

test('charts and images added to an output of automizer get unused names', async () => {
  // An output with gaps in its part numbering: the root template's charts
  // 1 and 2 are removed, the imported ones are numbered from 3.
  const first = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
    cleanup: true,
  });
  await first
    .loadRoot(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`SlideWithImages.pptx`, 'images')
    .addSlide('charts', 1)
    .addSlide('charts', 2)
    .addSlide('images', 1)
    .addSlide('images', 2)
    .write(`part-names-of-existing-parts-1.test.pptx`);

  const before = await readParts(`part-names-of-existing-parts-1.test.pptx`);

  const second = new Automizer({ templateDir, outputDir });
  await second
    .loadRoot(
      fs.readFileSync(`${outputDir}/part-names-of-existing-parts-1.test.pptx`),
    )
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`RootTemplateWithImages.pptx`, 'images')
    .load(`SVGImages.pptx`, 'svg')
    .addSlide('charts', 1)
    .addSlide('images', 1)
    .addSlide('svg', 1)
    .write(`part-names-of-existing-parts-2.test.pptx`);

  expectPartsKept(
    before,
    await readParts(`part-names-of-existing-parts-2.test.pptx`),
  );
});

test('diagrams added to a root template with gaps in its numbering get unused names', async () => {
  // SmartArt parts numbered 1, 2 and 4
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/SlideWithDiagrams.pptx`),
  );
  for (const kind of ['data', 'colors', 'layout', 'quickStyle', 'drawing']) {
    zip.file(
      `ppt/diagrams/${kind}4.xml`,
      await zip.file(`ppt/diagrams/${kind}3.xml`).async('nodebuffer'),
    );
    zip.remove(`ppt/diagrams/${kind}3.xml`);
  }
  for (const part of [
    '[Content_Types].xml',
    'ppt/slides/_rels/slide3.xml.rels',
  ]) {
    const xml = await zip.file(part).async('string');
    zip.file(part, xml.replace(/(diagrams\/[a-zA-Z]+)3\.xml/g, '$14.xml'));
  }
  const root = await zip.generateAsync({ type: 'nodebuffer' });

  const automizer = new Automizer({ templateDir, outputDir });
  await automizer
    .loadRoot(root)
    .load(`SlideWithDiagrams.pptx`, 'diagrams')
    .addSlide('diagrams', 1)
    .write(`part-names-of-existing-parts-diagrams.test.pptx`);

  expectPartsKept(
    await readParts(root),
    await readParts(`part-names-of-existing-parts-diagrams.test.pptx`),
  );
});

test('OLE objects of two slides are embedded as two parts', async () => {
  // Two slides with an OLE object each, related by the same rel id
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/threeSlidePres.pptx`),
  );
  for (const n of [1, 2]) {
    zip.file(`ppt/embeddings/oleObject${n}.bin`, `payload of slide ${n}`);
    const contentTypes = await zip.file('[Content_Types].xml').async('string');
    zip.file(
      '[Content_Types].xml',
      contentTypes.replace(
        '</Types>',
        `<Override PartName="/ppt/embeddings/oleObject${n}.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`,
      ),
    );
    const rels = `ppt/slides/_rels/slide${n}.xml.rels`;
    zip.file(
      rels,
      (await zip.file(rels).async('string')).replace(
        '</Relationships>',
        `<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject${n}.bin"/></Relationships>`,
      ),
    );
    const slide = `ppt/slides/slide${n}.xml`;
    zip.file(
      slide,
      (await zip.file(slide).async('string')).replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="90" name="Object"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
          '<p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>' +
          '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
          '<p:oleObj progId="Package" r:id="rId90"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree>',
      ),
    );
  }

  const automizer = new Automizer({ templateDir, outputDir });
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(await zip.generateAsync({ type: 'nodebuffer' }), 'ole')
    .addSlide('ole', 1)
    .addSlide('ole', 2)
    .write(`part-names-of-existing-parts-ole.test.pptx`);

  const output = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/part-names-of-existing-parts-ole.test.pptx`),
  );
  const payloads: string[] = [];
  for (const rels of output.file(/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/)) {
    const target = (await rels.async('string')).match(
      /relationships\/oleObject" Target="\.\.\/embeddings\/([^"]+)"/,
    );
    if (target) {
      payloads.push(
        await output.file(`ppt/embeddings/${target[1]}`).async('string'),
      );
    }
  }
  expect(payloads.sort()).toEqual(['payload of slide 1', 'payload of slide 2']);
});
