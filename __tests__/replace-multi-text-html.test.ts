import Automizer, { modify } from '../src/index';

test('create presentation, replace multi text from HTML string.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<html><body>\n' +
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
    '<p><strong>bold text</strong></p>\n' +
    '</body></html>\n';

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`modify-multi-text-html.test.pptx`);

  // expect(result.tables).toBe(2); // TODO: fixture for pptx-output
});
