import Automizer, { modify } from '../src/index';

test('create presentation, replace multi text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        modify.setMultiText([
          {
            paragraph: {
              bullet: true,
              level: 0,
              marginLeft: 41338,
              indent: -87325,
              alignment: 'l',
            },
            textRuns: [
              {
                text: 'test Hyperlink to github',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: 'CCCCCC',
                  },
                  hyperlink: {
                    type: 'external',
                    target: 'https://github.com/singerla/pptx-automizer',
                  },
                },
              },
            ],
          },
          {
            paragraph: {
              bullet: true,
              level: 1,
              marginLeft: 541338,
              indent: -187325,
            },
            textRuns: [
              {
                text: 'No Hyperlink',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: 'CCCCCC',
                  },
                },
              },
              {
                text: 'Internal Hyperlink to slide 1',
                style: {
                  size: 700,
                  color: {
                    type: 'srgbClr',
                    value: 'FF0000',
                  },
                  hyperlink: {
                    type: 'internal',
                    target: 1,
                  },
                },
              },
            ],
          },
        ]),
      );
    })
    .write(`modify-multi-text-hyperlink.test.pptx`);

  // expect(result.tables).toBe(2); // TODO: fixture for pptx-output
});
