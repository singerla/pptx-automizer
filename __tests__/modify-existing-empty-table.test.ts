import Automizer, { modify } from '../src/index';

test('Modify an existing table without content in any cell.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data1 = {
    body: [
      { label: 'item test r1', values: ['test1', 10, 16, 12, 11] },
      { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
      { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 3, (slide) => {
      slide.modifyElement('EmptyTable', [
        modify.setTable(data1),
        // modify.dump
      ]);
    })
    .write(`modify-existing-empty-table.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
