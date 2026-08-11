import Automizer from '../src/index';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import fs from 'fs';
import crypto from 'crypto';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

/**
 * Read all media files of an output .pptx and all relations pointing to them.
 */
const getMediaInfo = async (file: string) => {
  const archive = await JSZip.loadAsync(fs.readFileSync(`${outputDir}/${file}`));

  const mediaFiles = Object.keys(archive.files).filter(
    (name) => name.startsWith('ppt/media/') && !name.endsWith('/'),
  );

  const checksums: string[] = [];
  for (const mediaFile of mediaFiles) {
    const content = await archive.file(mediaFile).async('nodebuffer');
    checksums.push(crypto.createHash('sha1').update(content).digest('hex'));
  }

  /*
   * Resolve each image a slide actually displays: a:blip/@r:embed to the
   * relation it refers to, and the relation to the media part.
   * Unused relations left behind on a copied slide are a known issue and
   * not part of this - the shapes never point to them.
   */
  const images: { slide: string; target: string; exists: boolean }[] = [];
  const slides = Object.keys(archive.files).filter((name) =>
    name.match(/ppt\/slides\/slide\d+\.xml$/),
  );

  for (const slide of slides) {
    const slideXml = new DOMParser().parseFromString(
      await archive.file(slide).async('string'),
      'application/xml',
    );
    const relsXml = new DOMParser().parseFromString(
      await archive
        .file(slide.replace('slides/', 'slides/_rels/') + '.rels')
        .async('string'),
      'application/xml',
    );
    const relations = Array.from(relsXml.getElementsByTagName('Relationship'));

    for (const blip of Array.from(slideXml.getElementsByTagName('a:blip'))) {
      const rId = blip.getAttribute('r:embed');
      const target =
        relations
          .find((relation) => relation.getAttribute('Id') === rId)
          ?.getAttribute('Target') || '';

      images.push({
        slide,
        target,
        exists: archive.file(target.replace('../', 'ppt/')) !== null,
      });
    }
  }

  return { mediaFiles, checksums, images };
};

test('an image used on multiple slides is stored only once', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images');

  // Import the same two slides five times
  for (let i = 0; i < 5; i++) {
    pres.addSlide('images', 1);
    pres.addSlide('images', 2);
  }

  const result = await pres.write(`deduplicate-media.test.pptx`);

  const { mediaFiles, checksums, images } = await getMediaInfo(
    `deduplicate-media.test.pptx`,
  );

  // No media file is a copy of another one
  expect(new Set(checksums).size).toBe(mediaFiles.length);

  // Ten slides worth of images, each of them still displaying a picture
  expect(images.length).toBeGreaterThanOrEqual(10);
  expect(images.filter((image) => !image.exists)).toEqual([]);

  // The images of two slides, imported five times
  expect(mediaFiles.length).toBe(4);
  expect(result.images).toBe(4);
});

test('an imported image already contained in the root template is re-used', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
  });

  const pres = automizer
    // Root template and source template share their images
    .loadRoot(`RootTemplateWithImages.pptx`)
    .load(`RootTemplateWithImages.pptx`, 'same');

  const before = await getMediaCount(`RootTemplateWithImages.pptx`);

  pres.addSlide('same', 1);

  await pres.write(`deduplicate-media-root.test.pptx`);

  const { mediaFiles, checksums } = await getMediaInfo(
    `deduplicate-media-root.test.pptx`,
  );

  expect(new Set(checksums).size).toBe(mediaFiles.length);
  expect(mediaFiles.length).toBe(before);
});

const getMediaCount = async (template: string): Promise<number> => {
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${templateDir}/${template}`),
  );
  return Object.keys(archive.files).filter(
    (name) => name.startsWith('ppt/media/') && !name.endsWith('/'),
  ).length;
};
