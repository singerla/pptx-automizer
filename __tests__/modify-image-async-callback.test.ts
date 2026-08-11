import Automizer from '../src/automizer';
import { ModifyImageHelper } from '../src';
import { XmlElement } from '../src/types/xml-types';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const mediaDir = `${__dirname}/media`;

const readSlideFile = async (
  outputFile: string,
  file: string,
): Promise<string> => {
  const fileData = fs.readFileSync(path.join(outputDir, outputFile));
  const zip = await JSZip.loadAsync(fileData);
  const zipEntry = zip.file(file);

  expect(zipEntry).not.toBeNull();

  return zipEntry!.async('text');
};

test('Await an asynchronous ShapeModificationCallback before writing the slide', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
  });

  // A callback that modifies the element only after an awaited operation.
  // Without awaiting the callback, the slide would have been written to the
  // archive before "descr" was set.
  const setAltTextAsync = async (element: XmlElement): Promise<void> => {
    await new Promise((resolve) => setTimeout(resolve, 10));
    element
      .getElementsByTagName('p:cNvPr')
      .item(0)
      .setAttribute('descr', 'modified by async callback');
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images');

  const outputFile = `modify-image-async-callback.test.pptx`;

  await pres
    .addSlide('images', 1, (slide) => {
      slide.modifyElement('Grafik 5', [setAltTextAsync]);
    })
    .write(outputFile);

  const slideXml = await readSlideFile(outputFile, 'ppt/slides/slide2.xml');

  expect(slideXml).toContain('modified by async callback');
});

test('Await setRelationTargetCover to replace the image relation target', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    mediaDir,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .loadMedia(`test.png`)
    .load(`SlideWithImages.pptx`, 'images');

  const outputFile = `modify-image-async-cover.test.pptx`;

  await pres
    .addSlide('images', 1, (slide) => {
      slide.modifyElement('Grafik 5', [
        ModifyImageHelper.setRelationTargetCover('test.png', automizer),
      ]);
    })
    .write(outputFile);

  const slideRelsXml = await readSlideFile(
    outputFile,
    'ppt/slides/_rels/slide2.xml.rels',
  );

  // setRelationTargetCover updates the "Target" attribute after reading the
  // original image from the archive, which requires an awaited callback.
  expect(slideRelsXml).toContain('../media/test.png');
});
