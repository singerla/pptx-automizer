import Automizer, { modify, TableData } from '../src/index';

test('Add and modify an existing table, apply styles to cell.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data1: TableData = {
    body: [
      {
        values: ['test1', 10, 16, 12, 11],
        styles: [
          {
            color: {
              type: 'srgbClr',
              value: '00FF00',
            },
            background: {
              type: 'srgbClr',
              value: 'CCCCCC',
            },
            isItalics: true,
            isBold: true,
          },
        ],
      },
      { values: ['test2', 12, 18, 15, 12] },
      {
        values: ['test3', 14, 12, 11, 14],
        styles: [
          null,
          null,
          null,
          null,
          {
            color: {
              type: 'srgbClr',
              value: 'FF0000',
            },
            background: {
              type: 'srgbClr',
              value: '333333',
            },
            isItalics: true,
            isBold: true,
          },
        ],
      },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 3, (slide) => {
      slide.modifyElement('TableWithFormattedCells', [
        modify.setTable(data1),
        // modify.dump
      ]);
    })
    .write(`modify-existing-table-format-cells.test.pptx`);

  // Expect the first cell and the last cell to be formatted
  // expect(result.tables).toBe(2); // tbd
});
