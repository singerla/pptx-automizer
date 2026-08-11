import Automizer from '../src/index';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

type SlideParts = {
  /* r:id of each p:sldId in ppt/presentation.xml, in presentation order */
  sldIdLst: string[];
  /* All slide relationships of ppt/_rels/presentation.xml.rels */
  slideRels: { rId: string; target: string; exists: boolean }[];
};

/**
 * Read the slide list of a written .pptx from both places it is stored in:
 * the visible slides in `p:sldIdLst` and the slide relationships they point to.
 */
const getSlideParts = async (file: string): Promise<SlideParts> => {
  const archive = await JSZip.loadAsync(fs.readFileSync(`${outputDir}/${file}`));
  const parse = async (part: string) =>
    new DOMParser().parseFromString(
      await archive.file(part).async('string'),
      'application/xml',
    );

  const presentation = await parse('ppt/presentation.xml');
  const slideIds = presentation.getElementsByTagName('p:sldId');
  const sldIdLst = Array.from(slideIds).map((slideId) =>
    slideId.getAttribute('r:id'),
  );

  const rels = await parse('ppt/_rels/presentation.xml.rels');
  const relationships = Array.from(rels.getElementsByTagName('Relationship'));
  const slideRels = relationships
    .filter((relationship) =>
      relationship.getAttribute('Type').endsWith('/slide'),
    )
    .map((relationship) => {
      const target = relationship.getAttribute('Target');
      return {
        rId: relationship.getAttribute('Id'),
        target,
        exists: archive.file('ppt/' + target) !== null,
      };
    });

  return { sldIdLst, slideRels };
};

test('removeExistingSlides drops slide relations along with the slide ids', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
  });

  const pres = automizer
    // RootTemplate.pptx has one existing slide
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .addSlide('shapes', 2)
    .addSlide('shapes', 1);

  const result = await pres.write(`remove-existing-slides.test.pptx`);

  // The removed slide must not be counted any longer
  expect(result.slides).toBe(2);

  const { sldIdLst, slideRels } = await getSlideParts(
    `remove-existing-slides.test.pptx`,
  );

  expect(sldIdLst.length).toBe(2);

  // No leftover relation pointing to a removed slide, and every remaining
  // relation is used by p:sldIdLst and resolves to an existing part.
  expect(slideRels.map((rel) => rel.rId).sort()).toEqual([...sldIdLst].sort());
  expect(slideRels.filter((rel) => !rel.exists)).toEqual([]);
});

test('removeExistingSlides removes all slides of a multi slide root template', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
  });

  const pres = automizer
    // SlideTitles.pptx has three existing slides
    .loadRoot(`SlideTitles.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .addSlide('shapes', 2);

  const result = await pres.write(`remove-existing-slides-multi.test.pptx`);

  expect(result.slides).toBe(1);

  const { sldIdLst, slideRels } = await getSlideParts(
    `remove-existing-slides-multi.test.pptx`,
  );

  expect(sldIdLst.length).toBe(1);
  expect(slideRels.map((rel) => rel.rId)).toEqual(sldIdLst);
});

test('removeExistingSlides with cleanup leaves no relation to a deleted part', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
    // cleanup removes the slide parts that are no longer required
    cleanup: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .addSlide('shapes', 2);

  await pres.write(`remove-existing-slides-cleanup.test.pptx`);

  const { sldIdLst, slideRels } = await getSlideParts(
    `remove-existing-slides-cleanup.test.pptx`,
  );

  expect(sldIdLst.length).toBe(1);
  expect(slideRels.filter((rel) => !rel.exists)).toEqual([]);
});

test('re-importing a generated pptx yields the visible slides only', async () => {
  const automizer = new Automizer({
    templateDir: outputDir,
    outputDir,
  });

  const pres = automizer.load(
    `remove-existing-slides.test.pptx`,
    'generated',
  );

  const info = await pres.getInfo();
  const slides = info.slidesByTemplate('generated');

  // The two slides added above, not the removed template slide
  expect(slides.length).toBe(2);

  // Slide numbers stay file based, so they can be used with addSlide()
  expect(slides.map((slide) => slide.number)).toEqual([2, 3]);

  // Same for the slide numbers of the template itself
  const slideNumbers = await pres
    .getTemplate('generated')
    .getAllSlideNumbers();
  expect(slideNumbers).toEqual([2, 3]);
});
