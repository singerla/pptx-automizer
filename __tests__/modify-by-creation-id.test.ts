import Automizer, { modify } from '../src/index';

test('create presentation, add and modify an existing table by creation id.', async () => {
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

  const creationIds = await pres.setCreationIds()

  const result = await pres
    .addSlide('tables', 1950777067, (slide) => {
      slide.modifyElement(
        '{EFC74B4C-D832-409B-9CF4-73C1EFF132D8}',
        [modify.setTableData(data1)]);

      slide.addElement(
        'tables',
        1950777067,
        '{EFC74B4C-D832-409B-9CF4-73C1EFF132D8}',
        [modify.setTableData(data1)]);
    })
    .write(`modify-existing-table.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
