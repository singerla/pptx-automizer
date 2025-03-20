import Automizer, { modify } from '../src/index';

test('modify a table created with google slides, create text elements.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data1 = {
    body: [
      { values: ['', '', '', '', ''] },
      { values: ['', '', '', '', ''] },
      { values: ['', '', '', '', ''] },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`GS_Tables_Merged.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 4, (slide) => {
      slide.modifyElement('Google Shape;69;p16', [
        modify.setTable(data1),
        // modify.dump
      ]);
    })
    .write(`modify-existing-table-google-slides.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
