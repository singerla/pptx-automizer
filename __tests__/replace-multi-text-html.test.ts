import Automizer, { modify } from '../src/index';

test('create presentation, replace multi text from HTML string.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html = '<html><body><p>First Line 14pt</p>\n' +
    '<p><span style="font-size: 12px;">2nd line 12pt <strong>bold</strong> <em>italics</em></span></p>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 1 level 1</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 1 level 2</span></li>\n' +
    '</ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 1</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 2</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 3</span></li>\n' +
    '</ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 2</span></li>\n' +
    '<li><span style="font-size: 14px;"><ins>bullet</ins> <em>mixed</em> <strong><em>formatting</em></strong></span></li>\n' +
    '</ul>\n' +
    '</ul>\n' +
    '<p><span style="font-size: 14px;"><strong><em>Text </em></strong>after bullet list</span></p></body></html>\n'

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`modify-multi-text-html.test.pptx`);

  // expect(result.tables).toBe(2); // TODO: fixture for pptx-output
});
