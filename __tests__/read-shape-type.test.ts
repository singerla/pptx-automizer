import Automizer, { XmlElement, XmlHelper } from '../src/index';
import { ModifyShapeHelper } from '../src';

test('read shape type info', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`ShapeTypesCollection.pptx`, 'collection');

  pres.addSlide('collection', 1, async (slide) => {
    slide.modifyElement(
      'VecorShape (Box with arrow)',
      (element: XmlElement) => {
        const type = ModifyShapeHelper.getElementVisualType(element);
        expect(type).toBe('vectorShape');
      },
    );

    slide.modifyElement('Line (Arrow)', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('vectorLine');
    });

    slide.modifyElement('Table', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('table');
    });

    slide.modifyElement('Textfield', (element: XmlElement) => {
      // XmlHelper.dump(element)
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('textField');
    });

    slide.modifyElement('SmartArt (Diagram)', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('smartArt');
    });

    slide.modifyElement('Image', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('picture');
    });

    slide.modifyElement('Chart', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('chart');
    });

    slide.modifyElement('Pictogram', (element: XmlElement) => {
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('svgImage');
    });

    slide.modifyElement('SVG Image', (element: XmlElement) => {
      // XmlHelper.dump(element);
      const type = ModifyShapeHelper.getElementVisualType(element);
      expect(type).toBe('svgImage');
    });
  });

  await pres.write(`read-shape-type.test.pptx`);
});
