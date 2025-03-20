import Automizer, { modify } from '../src/index';

const bulletPoints = ['first line', 'second line', 'third line'];
const multiLevelBullet = [
  'top line 1',
  [
    'indent 1-1',
    'indent 1-2',
    [
      'indent 2-1',
      'indent 2-2',
    ],
  ],
  'top line 2',
];

test('create presentation, replace bullet list text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const result = await pres

    .addSlide('TextReplace.pptx', 2, (slide) => {
      slide.modifyElement(
        'replaceTextBullet1',
        modify.setBulletList(bulletPoints),
      );
    }).addSlide('TextReplace.pptx', 2, (slide) => {
      slide.modifyElement(
        'replaceTextBullet1',
        modify.setBulletList(multiLevelBullet),
      );
    })
    .write(`replace-bullet-text.test.pptx`);

   expect(result.slides).toBe(3);
});
