import Automizer, { modify } from './index';
import { vd } from './helper/general-helper';
import * as fs from 'fs';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  // Step 1: Create a pptx with images and a chart inside.
  // The chart is modified by pptx-automizer

  const automizer = new Automizer({
    templateDir,
    outputDir,
    verbosity: 2,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<p><span style="font-size: 24px;">Testing layouts and exporting them.</span></p>\n' +
    '<ul>\n' +
    '<li>level 1 - 1</li>\n' +
    '<li>level 1 - 2</li>\n' +
    '<ul>\n' +
    '<li>level 1-2-1 <em>italics</em></li>\n' +
    '</ul>\n' +
    '<li>level 1 - 3</li>\n' +
    '<ul>\n' +
    '<li>level 1 - 3 - 1</li>\n' +
    '</ul>\n' +
    '</ul>\n' +
    '<p>Testing testing testing</p>\n' +
    '<p><strong>bold text</strong></p>\n';

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        modify.setMultiText([
          {
            paragraph: {
              bullet: true,
              level: 1,
              marginLeft: 541338,
              indent: -187325,
            },
            textRuns: [
              {
                text: 'test ',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: 'CCCCCC',
                  },
                },
              },
              {
                text: 'test 2',
                style: {
                  size: 700,
                  color: {
                    type: 'srgbClr',
                    value: 'FF0000',
                  },
                },
              },
              {
                text: 'test 3',
                style: {
                  size: 1200,
                  color: {
                    type: 'srgbClr',
                    value: '00FF00',
                  },
                },
              },
            ],
          },
        ]),
      );
    })
    .write(`modify-multi-text.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
