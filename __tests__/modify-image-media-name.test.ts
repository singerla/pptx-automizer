import Automizer from '../src/automizer';
import { ModifyShapeHelper } from '../src';
import { expectXml } from './helpers/expect-xml';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

/**
 * automizer names the media files it copies `media<n>.<ext>`, PowerPoint
 * names them `image<n>.<ext>`. A picture must be found by its relationship
 * either way - e.g. on a slide of an output of automizer that is loaded as
 * a template again.
 */
test('modify, add and remove a picture whose media file is not named image<n>', async () => {
  const first = new Automizer({ templateDir, outputDir });
  await first
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images')
    .addSlide('images', 1)
    .write(`modify-image-media-name-1.test.pptx`);

  // The copied picture refers to ppt/media/media<n>.jpg
  const rels = await expectXml(
    `modify-image-media-name-1.test.pptx`,
    'ppt/slides/_rels/slide2.xml.rels',
  );
  rels.toHaveAttribute('Relationship', 'Target', '../media/media1.jpg');

  const second = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
  });
  const pres = second
    .loadRoot(`RootTemplate.pptx`)
    .load(
      fs.readFileSync(`${outputDir}/modify-image-media-name-1.test.pptx`),
      'output',
    )
    .load(`SlideWithImages.pptx`, 'images');

  pres.addSlide('output', 2, (slide) => {
    slide.modifyElement('Grafik 5', [
      ModifyShapeHelper.setPosition({ x: 1000000 }),
    ]);
  });
  pres.addSlide('images', 1, (slide) => {
    slide.addElement('output', 2, 'Grafik 5');
  });
  pres.addSlide('output', 2, (slide) => {
    slide.removeElement('Grafik 5');
  });

  const result = await pres.write(`modify-image-media-name-2.test.pptx`);
  expect(result.slides).toBe(3);

  const output = `modify-image-media-name-2.test.pptx`;
  (await expectXml(output, 'ppt/slides/slide2.xml'))
    .toContainElementTimes('p:pic', 1)
    .toHaveAttribute('a:off', 'x', '1000000');
  (await expectXml(output, 'ppt/slides/slide3.xml')).toContainElementTimes(
    'p:pic',
    2,
  );
  (await expectXml(output, 'ppt/slides/slide4.xml')).toContainElementTimes(
    'p:pic',
    0,
  );
});
