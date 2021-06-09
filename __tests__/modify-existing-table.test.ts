import Automizer, { modify } from '../src/index';

test('create presentation, add and modify an existing table.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data1 = {
    body: [
      { label: 'item test r1', values: ['test1', 10, 16, 12] },
      { label: 'item test r2', values: ['test2', 12, 18, 15] },
      { label: 'item test r3', values: ['test3', 14, 12, 11] },
      // { label: 'item test r4', values: ['test4', 14, 12, 11] },
      // { label: 'item test r5', values: ['test5', 14, 12, 11] },
      // { label: 'item test r6', values: ['test6', 999, 12, 11] },
      // { label: 'item test r6', values: ['test7', 999, 12, 11] },
      // { label: 'item test r6', values: ['test8', 999, 12, 11] },
      // { label: 'item test r6', values: ['test9', 999, 12, 11] },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableWithHeader', [modify.setTableData(data1)]);
    })
    .write(`modify-existing-table.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
