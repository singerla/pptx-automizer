import Automizer, { modify } from '../src/index';


test('create presentation, replace multi text.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.setMultiText([
        {
          paragraph: {
            bullet: true,
            level: 0,
            marginLeft: 41338,
            indent: -87325,
            alignment: 'left'
          },
          textRuns: [
            {
              text: 'test 0',
              style: {
                color: {
                  type: 'srgbClr',
                  value: 'CCCCCC'
                }
              }
            },
          ]
        },
        {
          paragraph: {
            bullet: true,
            level: 1,
            marginLeft: 541338,
            indent: -187325,
            alignment: 'left'
          },
          textRuns: [
            {
              text: 'test ',
              style: {
                color: {
                  type: 'srgbClr',
                  value: 'CCCCCC'
                }
              }
            },
            {
              text: 'test 2',
              style: {
                size: 700,
                color: {
                  type: 'srgbClr',
                  value: 'FF0000'
                }
              }
            },
            {
              text: 'test 3',
              style: {
                size: 1200,
                color: {
                  type: 'srgbClr',
                  value: '00FF00'
                }
              }
            }
          ]
        },
        {
          paragraph: {
            alignment: 'right'
          },
          textRuns: [
            {
              text: 'aligned Center',
              style: {
                color: {
                  type: 'srgbClr',
                  value: '00FF00'
                }
              }
            },
          ]
        },
      ]));
    })
    .write(`modify-multi-text.test.pptx`);

  // expect(result.tables).toBe(2); // TODO: fixture for pptx-output
});
