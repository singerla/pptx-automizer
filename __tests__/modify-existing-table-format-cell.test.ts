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
            size: 1200,
          },
        ],
      },
      {
        values: ['test2', 12, 18, 15, 12],
        styles: [
          null,
          null,
          null,
          null,
          {
            // If you want to style a cell border, you
            // need to style adjacent borders as well:
            border: [
              {
                // This is required to complete top border
                // of adjacent cell in row below:
                tag: 'lnB',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
                },
              },
            ],
          },
        ],
      },
      {
        values: ['test3', 14, 12, 11, 14],
        styles: [
          null,
          null,
          null,
          {
            border: [
              {
                tag: 'lnR',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
                },
              },
            ],
          },
          {
            color: {
              type: 'srgbClr',
              value: 'FF0000',
            },
            background: {
              type: 'srgbClr',
              value: 'ffffff',
            },
            isItalics: true,
            isBold: true,
            size: 600,
            border: [
              {
                // This will only work in case you style
                // adjacent cell in row above with 'lnB':
                tag: 'lnT',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
                },
              },
              {
                tag: 'lnB',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
                },
              },
              {
                tag: 'lnL',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
                },
              },
              {
                tag: 'lnR',
                type: 'solid',
                weight: 5000,
                color: {
                  type: 'srgbClr',
                  value: '00FF00',
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
