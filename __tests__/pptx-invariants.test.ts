import fs from 'fs';
import JSZip from 'jszip';
import { checkPptxInvariants } from './helpers/pptx-invariants';

/**
 * Self-test for the Tier-1 invariant checker: a clean template passes, and
 * each corruption class is detected when introduced artificially. Guards the
 * guard — a checker that silently stops detecting would turn the whole tier
 * into a no-op.
 */

const templatePath = `${__dirname}/pptx-templates/SlideWithImages.pptx`;

const loadTemplate = async (): Promise<JSZip> =>
  JSZip.loadAsync(fs.readFileSync(templatePath));

const toBuffer = (archive: JSZip): Promise<Buffer> =>
  archive.generateAsync({ type: 'nodebuffer' });

test('a clean template produces no errors', async () => {
  const { errors } = await checkPptxInvariants(
    fs.readFileSync(templatePath),
  );
  expect(errors).toEqual([]);
});

test('detects a referenced relationship with a missing target part', async () => {
  const archive = await loadTemplate();
  const mediaPart = Object.keys(archive.files).find((name) =>
    name.startsWith('ppt/media/'),
  );
  archive.remove(mediaPart);

  const { errors } = await checkPptxInvariants(await toBuffer(archive));

  expect(
    errors.some((error) => error.includes(`missing part ${mediaPart}`)),
  ).toBe(true);
});

test('detects an r:embed without a relationship entry', async () => {
  const archive = await loadTemplate();
  const slideXml = await archive.file('ppt/slides/slide1.xml').async('text');
  archive.file(
    'ppt/slides/slide1.xml',
    slideXml.replace(/r:embed="[^"]+"/, 'r:embed="rId999"'),
  );

  const { errors } = await checkPptxInvariants(await toBuffer(archive));

  expect(errors.some((error) => error.includes('rId999'))).toBe(true);
});

test('detects a part missing from [Content_Types].xml', async () => {
  const archive = await loadTemplate();
  archive.file('ppt/unregistered.bin', 'not registered anywhere');

  const { errors } = await checkPptxInvariants(await toBuffer(archive));

  expect(
    errors.some((error) =>
      error.includes('ppt/unregistered.bin: not covered'),
    ),
  ).toBe(true);
});

test('detects a slide list entry without a slide part', async () => {
  const archive = await loadTemplate();
  const slideCount = Object.keys(archive.files).filter((name) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(name),
  ).length;
  expect(slideCount).toBeGreaterThan(0);
  archive.remove(`ppt/slides/slide${slideCount}.xml`);

  const { errors } = await checkPptxInvariants(await toBuffer(archive));

  expect(
    errors.some((error) =>
      error.includes(`slide list entry ppt/slides/slide${slideCount}.xml`),
    ),
  ).toBe(true);
});

test('detects a part that is not well-formed XML', async () => {
  const archive = await loadTemplate();
  archive.file('ppt/slides/slide1.xml', '<p:sld><unclosed></p:sld>');

  const { errors } = await checkPptxInvariants(await toBuffer(archive));

  expect(
    errors.some((error) =>
      error.includes('ppt/slides/slide1.xml: not well-formed XML'),
    ),
  ).toBe(true);
});

test('reports stale unreferenced relationships as known issues, not errors', async () => {
  const archive = await loadTemplate();
  const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const rels = await archive.file(relsPath).async('text');
  archive.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdStale" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/gone.png"/></Relationships>',
    ),
  );

  const { errors, knownIssues } = await checkPptxInvariants(
    await toBuffer(archive),
  );

  expect(errors).toEqual([]);
  expect(knownIssues.some((issue) => issue.includes('rIdStale'))).toBe(true);
});
