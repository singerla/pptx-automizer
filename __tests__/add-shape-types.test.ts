import Automizer, { XmlElement, XmlHelper } from '../src/index';
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

    slide.addElement('collection', 1, 'Textfield', (element: XmlElement) => {
      // XmlHelper.dump(element)
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('textField');
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
  });

  await pres.write(`add-shape-types.test.pptx`);
});
