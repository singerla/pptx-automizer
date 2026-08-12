/**
 * Golden deck `tables` (Tier 3): setTable/setTableData growing and shrinking
 * rows+columns, cell styles and borders — mirrors modify-existing-table, but
 * judged by rendered pixels instead of XML.
 */
import Automizer, { modify, TableRow, TableRowStyle } from '../../src/index';
import { expectDeckMatchesBaselines } from './helpers/golden-deck';

test('golden deck: tables', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../pptx-templates`,
    outputDir: `${__dirname}/../pptx-output`,
  });

  // Grown table with a fully bordered first row
  const borderedGrow = {
    body: [
      <TableRow>{
        label: 'item test r1',
        values: ['test1', 10, 16, 12, 11],
        styles: [
          {
            border: (['lnB', 'lnR', 'lnL', 'lnT'] as const).map((tag) => ({
              tag,
              weight: 18500,
              type: 'sysDot',
              color: { type: 'srgbClr', value: 'aacc00' },
            })),
          },
        ],
      },
      { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
      { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
    ],
  };

  // Nine rows: growth well past the template's row count
  const nineRows = {
    body: Array.from({ length: 9 }, (_, i) => ({
      label: `item test r${i + 1}`,
      values: [`test${i + 1}`, 990 + i, 10 + i, 12],
    })),
  };

  // Shrunk table with a styled cell, height/width re-fitted
  const shrunkStyled = {
    body: [
      <TableRow>{
        label: 'item test r1',
        values: ['test1', 10, 16],
        styles: [
          null,
          <TableRowStyle>{
            color: { type: 'srgbClr', value: 'cccccc' },
            size: 1400,
          },
        ],
      },
      { label: 'item test r2', values: ['test2', 12, 18] },
      { label: 'item test r3', values: ['test3', 14, 13] },
    ],
  };

  const outputFile = 'visual-tables.pptx';
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables')
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableDefault', [modify.setTable(borderedGrow)]);
      slide.modifyElement('TableWithLabels', [modify.setTable(nineRows)]);
      slide.modifyElement('TableWithHeader', [
        modify.setTableData(shrunkStyled),
        modify.adjustHeight(shrunkStyled),
        modify.adjustWidth(shrunkStyled),
      ]);
    })
    .addSlide('tables', 2, (slide) => {
      slide.modifyElement('LabelsVertical', [
        modify.setTable({
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
                      color: { type: 'srgbClr', value: 'aacc00' },
                    },
                  ],
                },
              ],
            },
          ],
        }),
      ]);
    })
    .write(outputFile);

  await expectDeckMatchesBaselines('tables', outputFile, 3);
});
