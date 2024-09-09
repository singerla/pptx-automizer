import Automizer, { CmToDxa, modify, TableData } from '../src/index';

test('update row hight/column width in a nested table', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables');

  await pres
    .addSlide('tables', 4, async (slide) => {
      const tableData: TableData = {
        body: [
          { values: ['top left', 123, 345, 'subsub3-1', 'subsub3-2', 'Last'] },
          { values: [undefined, 't1', 't2', 't3', 't3', ''] },
          { values: ['label 0', 1, 2, 3, 3, 'l0'] },
          { values: ['label 1', 123, 345, 4563, 4671, 'l1'] },
          { values: ['label 2', 123, 345, 4562, 4672] },
          { values: ['label 3', 123, 345, 4561, 4673, 'l3'] },
          { values: [undefined, 'Foo', 'ter', 4564, 4674, ''] },
        ],
      };

      slide.modifyElement('NestedTable3', [
        modify.setTable(tableData, {
          expand: [
            {
              mode: 'row',
              tag: '{{each:row}}',
              count: 3,
            },
            {
              mode: 'column',
              tag: '{{each:subSub3}}',
              count: 1,
            },
          ],
        }),
        // Update first column width to 8cm
        modify.updateColumnWidth(0, CmToDxa(8)),
        // Update last row's height to 3cm
        modify.updateRowHeight(6, CmToDxa(3)),
        // Will also work on non-nested tables
      ]);
    })
    .write(`modify-nested-table-grid-size.test.pptx`);

  // expect(data.length).toBe(12);
});
