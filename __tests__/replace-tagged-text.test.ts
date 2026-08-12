import Automizer, { modify } from '../src/index';
import { expectXml } from './helpers/expect-xml';

const bulletPoints = ['first line', 'second line', 'third line'].join(`
`);

test('create presentation, replace tagged text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.setText('Test'));

      slide.modifyElement(
        'replaceText',
        modify.replaceText(
          [
            {
              replace: 'replace',
              by: {
                text: 'Apples',
              },
            },
            {
              replace: 'by',
              by: {
                text: 'Bananas',
              },
            },
            {
              replace: 'replacement',
              by: [
                {
                  text: 'Really!',
                  style: {
                    size: 10000,
                    color: {
                      type: 'srgbClr',
                      value: 'ccaa4f',
                    },
                  },
                },
                {
                  text: 'Fine!',
                  style: {
                    size: 10000,
                    color: {
                      type: 'schemeClr',
                      value: 'accent2',
                    },
                  },
                },
              ],
            },
          ],
          {
            openingTag: '{{',
            closingTag: '}}',
          },
        ),
      );
    })
    .addSlide('TextReplace.pptx', 2, (slide) => {
      slide.modifyElement(
        'replaceTextBullet1',
        modify.replaceText(
          [
            {
              replace: 'bullet1',
              by: {
                text: bulletPoints,
              },
            },
            {
              replace: 'bullet2',
              by: {
                text: bulletPoints,
              },
            },
          ],
          {
            openingTag: '{{',
            closingTag: '}}',
          },
        ),
      );
    })
    .write(`replace-tagged-text.test.pptx`);

  // Tier 0: replacements and their styles arrived in the slide XML
  const slide2 = await expectXml(
    'replace-tagged-text.test.pptx',
    'ppt/slides/slide2.xml',
  );
  slide2.toContainElement('a:t', 'Test');
  expect(slide2.raw()).toMatch(/Apples/);
  expect(slide2.raw()).toMatch(/Bananas/);
  expect(slide2.raw()).not.toMatch(/\{\{/);
  slide2
    .toHaveAttribute('a:srgbClr', 'val', 'CCAA4F')
    .toHaveAttribute('a:schemeClr', 'val', 'accent2');

  const slide3 = await expectXml(
    'replace-tagged-text.test.pptx',
    'ppt/slides/slide3.xml',
  );
  expect(slide3.raw()).toMatch(/second line/);
});
