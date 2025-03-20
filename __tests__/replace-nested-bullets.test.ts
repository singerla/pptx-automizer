import Automizer, { modify } from '../src/index';

const bulletPoints = ['first line', 'second line', 'third line'].join(`\n`);
const bulletPoints2 = ['first line-2', 'second line-2'].join(`\n`);

test('create presentation, replace tagged and untagged nested bulleted text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`TextReplaceBullets.pptx`);

  await pres
    .addSlide('TextReplaceBullets.pptx', 1, (slide) => {
      slide.modifyElement(
        '@AutomateBullets',
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
            {
              replace: 'bullet2-1',
              by: {
                text: bulletPoints2,
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
    .addSlide('TextReplaceBullets.pptx', 2, (slide) => {
      slide.modifyElement(
        '@AutomateNestedBullets',
        modify.setBulletList([
          'Test1',
          'Test 2',
          ['Test 3-1', 'Test 3-2', ['Test 3-3-1', 'Test 3-3-2']],
          'Test 4',
        ]),
      );
    })
    .write(`replace-nested-bullets.test.pptx`);
});
