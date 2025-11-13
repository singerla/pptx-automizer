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
                text: 'Internal Hyperlink',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: 'CCCCCC',
                  },
                  hyperlink: {
                    isInternal: true,
                    target: 1
                  }
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
          {
            paragraph: {
              alignment: 'r',
            },
            textRuns: [
              {
                text: 'aligned Right',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: '00FF00',
                  },
                },
              },
            ],
          },
          {
            paragraph: {
              alignment: 'ctr',
            },
            textRuns: [
              {
                text: 'aligned Center',
                style: {
                  color: {
                    type: 'srgbClr',
                    value: '0000FF',
                  },
                  hyperlink: {
                    target: 'https://github.com',
                  }
                },
              },
            ],
          },
        ]),
      );

      // Generate a textfield with pptxGenJS and modify its contents with setMultiText
      // This way, you can work around some pptxGenJS restrictions.
      slide.generate((iSlide) => {
        iSlide.addText(
          'test',
          {
            x: 10,
            h: 3,
            w: 3,
            valign: 'top',
            fontSize: 14,
          },
          [
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
                    text: 'test 123',
                    style: {
                      color: {
                        type: 'srgbClr',
                        value: 'CCCCCC',
                      },
                    },
                  },
                ],
              },
            ]),
          ],
        );
      }, 'my object');
    })
    .write(`modify-multi-text.test.pptx`);

  // expect(result.tables).toBe(2); // TODO: fixture for pptx-output
});
