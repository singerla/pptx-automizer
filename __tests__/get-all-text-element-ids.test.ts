import Automizer, { modify } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('create presentation, modify text elements using getAllTextElementIds.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`twoTextElementPres.pptx`);

  // Add a slide from the template and use getAllTextElementIds inside the callback
  pres.addSlide('twoTextElementPres.pptx', 1, async (slide) => {
    // Use the getAllTextElementIds method to get all text element IDs in the slide
    const elementIds = await slide.getAllTextElementIds();
    expect(elementIds.length).toEqual(2); // Assert that there are 2 element IDs

    // Loop through the element IDs and modify the text
    for (const elementId of elementIds) {
      slide.modifyElement(
        elementId,
        modify.replaceText(
          [
            {
              replace: 'placeholder',
              by: {
                text: 'New Text',
              },
            },
            {
              replace: 'placeholder2',
              by: {
                text: 'New Text 2',
              },
            },
          ],
          {
            openingTag: '{',
            closingTag: '}',
          },
        ),
      );
    }
  });

  const result = await pres.write(`get-all-text-element-ids.test.pptx`);
});
