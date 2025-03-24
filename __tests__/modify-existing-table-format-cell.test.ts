import Automizer, {
  modify,
  TableData,
  XmlElement,
  XmlHelper,
} from '../src/index';
import ModifyColorHelper from '../src/helper/modify-color-helper';

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
          {
            background: {
              type: 'srgbClr',
              value: 'FFFFFF',
            },
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

  // Retrieve all text contents from a given table cell:
  const parseCellText = (cell: XmlElement) => {
    const texts = cell.getElementsByTagName('a:t');
    const textFragments = [];
    for (let t = 0; t < texts.length; t++) {
      textFragments.push(texts.item(t).textContent);
    }
    return textFragments.join('');
  };

  // Iterate all rows and columns and appy a callback on each table cell:
  const modifyTableCell = (
    element: XmlElement,
    callback: (cell: XmlElement) => void,
  ) => {
    const rows = element.getElementsByTagName('a:tr');
    XmlHelper.modifyCollection(rows, (row) => {
      const columns = row.getElementsByTagName('a:tc');
      XmlHelper.modifyCollection(columns, (cell: XmlElement) => {
        callback(cell);
      });
    });
  };

  const result = await pres
    .addSlide('tables', 3, (slide) => {
      slide.modifyElement('TableWithEmptyCells', [
        // First, set up table with data & style
        modify.setTable(data1),

        // Show the table XML after setTable modification was done:
        // modify.dump

        // Apply a custom modifier if you require additional formatting:
        (element: XmlElement) => {
          modifyTableCell(element, (cell: XmlElement) => {
            // Get text from all text elements in current cell:
            const text = parseCellText(cell);

            // Do your checks:
            if (text === 'test1' || text === '11') {
              // Log cell XML to console:
              // XmlHelper.dump(cell);

              // Please note: this will only work on cells that do already
              // have a background color.
              ModifyColorHelper.solidFill(
                {
                  type: 'srgbClr',
                  value: '#FF3300',
                },
                'last',
              )(cell);
            }
          });
        },
      ]);
    })
    .write(`modify-existing-table-format-cells.test.pptx`);

  // Expect the first cell and the last cell to be formatted
  // expect(result.tables).toBe(2); // tbd
});
