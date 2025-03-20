import Automizer, { modify, TableData } from '../src/index';

test('modify a nested table with tags', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables');

  await pres
    .addSlide('tables', 3, async (slide) => {
      const tableData: TableData = {
        body: [
          {
            values: [
              'top left',
              'sub-1',
              null,
              null,
              'sub-2',
              null,
              null,
              'Last',
            ],
            styles: [
              {
                border: [
                  {
                    tag: 'lnB',
                    weight: 35000,
                    color: {
                      type: 'srgbClr',
                      value: 'aacc00',
                    },
                  },
                ],
              },
              {
                border: [
                  {
                    tag: 'lnB',
                    weight: 8500,
                    type: 'sysDot',
                    color: {
                      type: 'srgbClr',
                      value: 'aacc00',
                    },
                  },
                  {
                    tag: 'lnR',
                    weight: 8500,
                    type: 'sysDot',
                    color: {
                      type: 'srgbClr',
                      value: 'aacc00',
                    },
                  },
                ],
              },
            ],
          },
          { values: [undefined, 't1', 't2', 't3', 't3', 't3', 't3', ''] },
          { values: ['label 0', 1, 2, 3, 3, 't3', 't3', 'l0'] },
          { values: ['label 1', 123, 345, 4563, 4671, 't3', 't3', 'l1'] },
          { values: ['label 2', 123, 345, 4562, 4672] },
          { values: ['label 3', 123, 345, 4561, 4673, 't3', 't3', 'l3'] },
          { values: ['', 'Foo', 'ter', 4564, 'foo2', 't3', 't3', ''] },
        ],
      };

      slide.modifyElement(
        'NestedTable3',
        modify.setTable(tableData, {
          expand: [
            {
              mode: 'row',
              tag: '{{each:row}}',
              count: 3,
            },
            {
              mode: 'column',
              tag: '{{each:subSub2}}',
              count: 1,
            },
            {
              mode: 'column',
              tag: '{{each:sub}}',
              count: 1,
            },
          ],
        }),
      );
    })
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

      slide.modifyElement(
        'NestedTable3',
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
      );
    })
    .write(`modify-nested-table.test.pptx`);

  // expect(data.length).toBe(12);
});
