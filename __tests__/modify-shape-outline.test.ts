import Automizer, { modify, PtToEmu, XmlElement } from '../src/index';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import fs from 'fs';

/**
 * Read a written .pptx and return the <a:ln> of a shape, serialized.
 * Looks up the shape by name across all slides of the output.
 */
const getOutlineXml = async (
  file: string,
  shapeName: string,
): Promise<string> => {
  const archive = await JSZip.loadAsync(fs.readFileSync(file));
  const slides = Object.keys(archive.files).filter((name) =>
    name.match(/ppt\/slides\/slide\d+\.xml$/),
  );

  for (const slide of slides) {
    const xml = await archive.file(slide).async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const shapes = doc.getElementsByTagName('p:sp');

    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes.item(i) as unknown as XmlElement;
      const nvPr = shape.getElementsByTagName('p:cNvPr')[0];

      if (nvPr?.getAttribute('name') !== shapeName) continue;

      const spPr = shape.getElementsByTagName('p:spPr')[0];
      const lines = spPr.getElementsByTagName('a:ln');

      // There must never be more than one outline per shape
      expect(lines.length).toBeLessThanOrEqual(1);

      if (!lines[0]) return '';

      // Serializing a fragment repeats the namespace declaration of the
      // document root on the outermost tag - not part of the actual file.
      return new XMLSerializer()
        .serializeToString(lines[0] as any)
        .replace(/ xmlns:a="[^"]*"/, '');
    }
  }

  throw new Error(`Shape "${shapeName}" not found in ${file}`);
};

test('modify shape outline: weight, color and dash style', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const outputFile = `modify-shape-outline.test.pptx`;

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`EmptySlidePlaceholders.pptx`, 'placeholders');

  await pres
    // All shapes on this slide come with <a:ln><a:noFill/></a:ln>
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement(
        'Cloud',
        modify.setOutline({
          weight: PtToEmu(2),
          color: { type: 'srgbClr', value: 'FF0000' },
        }),
      );
      slide.modifyElement('Arrow', modify.setOutline({ weight: PtToEmu(4.5) }));
      slide.modifyElement(
        'Drum',
        modify.setOutline({
          type: 'sysDash',
          color: { type: 'schemeClr', value: 'accent2' },
        }),
      );
      // 'Star' is left untouched on purpose
    })
    // 'Textplatzhalter 5' has an empty <p:spPr/>: no outline to modify yet
    .addSlide('placeholders', 1, (slide) => {
      slide.modifyElement(
        'Textplatzhalter 5',
        modify.setOutline({
          weight: PtToEmu(1),
          type: 'dash',
          color: { type: 'srgbClr', value: '00FF00' },
        }),
      );
    })
    .write(outputFile);

  const output = `${__dirname}/pptx-output/${outputFile}`;

  // The written deck contains the two added slides
  const archive = await JSZip.loadAsync(fs.readFileSync(output));
  const presentation = await archive.file('ppt/presentation.xml').async('string');
  const slideIds = new DOMParser()
    .parseFromString(presentation, 'application/xml')
    .getElementsByTagName('p:sldId');
  expect(slideIds.length).toBe(2);

  // Existing outline: a:noFill is replaced by the given color
  expect(await getOutlineXml(output, 'Cloud')).toBe(
    '<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
  );

  // Weight only: the template's a:noFill is preserved (invisible outline)
  expect(await getOutlineXml(output, 'Arrow')).toBe(
    '<a:ln w="57150"><a:noFill/></a:ln>',
  );

  // Existing solidFill is replaced, a:prstDash inserted after the fill
  expect(await getOutlineXml(output, 'Drum')).toBe(
    '<a:ln><a:solidFill><a:schemeClr val="accent2"/></a:solidFill>' +
      '<a:prstDash val="sysDash"/></a:ln>',
  );

  // Untouched shape keeps its template outline
  expect(await getOutlineXml(output, 'Star')).toBe('<a:ln><a:noFill/></a:ln>');

  // Missing outline is created in schema order (fill, then dash)
  expect(await getOutlineXml(output, 'Textplatzhalter 5')).toBe(
    '<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>' +
      '<a:prstDash val="dash"/></a:ln>',
  );
});
