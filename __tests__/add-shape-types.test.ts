import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import Automizer, { XmlElement } from '../src/index';
import { ModifyShapeHelper } from '../src';

test('add all implemented shape types to an empty slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`ShapeTypesCollection.pptx`, 'collection');

  pres.addSlide('empty', 1, async (slide) => {
    slide.addElement(
      'collection',
      1,
      'VecorShape (Box with arrow)',
      (element: XmlElement) => {
        const type = ModifyShapeHelper.getElementVisualType(element);
        expect(type).toBe('vectorShape');
      },
    );

    slide.addElement('collection', 1, 'Line (Arrow)', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('vectorLine');
    });

    slide.addElement('collection', 1, 'Table', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('table');
    });

    slide.addElement('collection', 1, 'Textfield (Native)', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('textBox');
    });

    slide.addElement('collection', 1, 'Textfield', (element: XmlElement) => {
      // XmlHelper.dump(element)
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('textBox');
    });

    slide.addElement(
      'collection',
      1,
      'SmartArt (Diagram)',
      (element: XmlElement) => {
        const type = ModifyShapeHelper.getElementVisualType(element);
        expect(type).toBe('smartArt');
      },
    );

    slide.addElement('collection', 1, 'Image', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('picture');
    });

    slide.addElement('collection', 1, 'Chart', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('chart');
    });

    slide.addElement('collection', 1, 'Pictogram', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('svgImage');
    });

    slide.addElement('collection', 1, 'SVG Image', (element: XmlElement) => {
      // XmlHelper.dump(element);
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('svgImage');
    });

    slide.addElement('collection', 1, 'Image filled rectangle', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('svgImage');
    });
  });

  const result = await pres.write(`add-shape-types.test.pptx`);
  expect(result.slides).toBe(2);

  // An image filling a shape is not a p:pic, its media files and relations
  // need to be imported nevertheless.
  const zip = await JSZip.loadAsync(
    fs.readFileSync(path.join(`${__dirname}/pptx-output`, `add-shape-types.test.pptx`)),
  );
  const slideXml = await zip.file('ppt/slides/slide2.xml')!.async('text');
  const relsXml = await zip
    .file('ppt/slides/_rels/slide2.xml.rels')!
    .async('text');

  const embeddedRIds = Array.from(slideXml.matchAll(/r:embed="([^"]+)"/g)).map(
    (match) => match[1],
  );
  expect(embeddedRIds.length).toBe(7);

  embeddedRIds.forEach((rId) => {
    const rel = relsXml.match(new RegExp(`<Relationship Id="${rId}"[^>]+>`));
    expect(rel).not.toBeNull();

    // Every referenced media file made it into the archive.
    const target = rel![0].match(/Target="\.\.\/media\/([^"]+)"/);
    expect(target).not.toBeNull();
    expect(zip.file(`ppt/media/${target![1]}`)).not.toBeNull();
  });
});
