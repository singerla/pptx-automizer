import Automizer, { modify, XmlHelper } from '../src/index';
import { TableRow, TableRowStyle } from '../src/index';

test('create presentation, add and modify an existing table.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data1 = {
    body: [
      <TableRow>{
        label: 'item test r1',
        values: ['test1', 10, 16, 12, 11],
        styles: [
          {
            border: [
              {
                tag: 'lnB',
                weight: 18500,
                type: 'sysDot',
                color: {
                  type: 'srgbClr',
                  value: 'aacc00',
                },
              },

              {
                tag: 'lnR',
                weight: 18500,
                type: 'sysDot',
                color: {
                  type: 'srgbClr',
                  value: 'aacc00',
                },
              },
              {
                tag: 'lnL',
                weight: 18500,
                type: 'sysDot',
                color: {
                  type: 'srgbClr',
                  value: 'aacc00',
                },
              },

              {
                tag: 'lnT',
                weight: 18500,
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
      { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
      { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
    ],
  };

  const data2 = {
    body: [
      { label: 'item test r1', values: ['test1', 10, 16, 12] },
      { label: 'item test r2', values: ['test2', 12, 18, 15] },
      { label: 'item test r3', values: ['test3', 14, 12, 11] },
      { label: 'item test r4', values: ['test4', 14, 12, 18] },
      { label: 'item test r5', values: ['test5', 14, 13, 15] },
      { label: 'item test r6', values: ['test6', 999, 14, 14] },
      { label: 'item test r7', values: ['test7', 998, 15, 13] },
      { label: 'item test r8', values: ['test8', 997, 16, 19] },
      { label: 'item test r9', values: ['test9', 996, 17, 18] },
    ],
  };

  const data3 = {
    body: [
      <TableRow>{
        label: 'item test r1',
        values: ['test1', 10, 16],
        styles: [
          null,
          <TableRowStyle>{
            color: {
              type: 'srgbClr',
              value: 'cccccc',
            },
            size: 1400,
          },
        ],
      },
      { label: 'item test r2', values: ['test2', 12, 18] },
      {
        label: 'item test r3',
        values: ['test3', 14, 13],
      },
    ],
  };

  const data4 = {
    body: [
      <TableRow>{
        label: 'item test r1',
        values: ['test1', 12],
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
                tag: 'lnT',
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
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableDefault', [modify.setTable(data1)]);

      slide.modifyElement('TableWithLabels', [
        modify.setTable(data2),
        // modify.dump
      ]);

      slide.modifyElement('TableWithHeader', [
        modify.setTableData(data3),
        modify.adjustHeight(data3),
        modify.adjustWidth(data3),
      ]);
    })
    .addSlide('tables', 2, (slide) => {
      slide.modifyElement('LabelsVertical', [modify.setTable(data4)]);
    })
    .write(`modify-existing-table.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
