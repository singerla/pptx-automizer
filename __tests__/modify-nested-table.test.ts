import Automizer, { XmlDocument, XmlElement, XmlHelper } from '../src/index';
import { ChartModificationCallback } from '../src/types/types';
import { vd } from '../src/helper/general-helper';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables');
  const data = [];

  const readTableData =
    (data: any): ChartModificationCallback =>
    (element: XmlElement): void => {
      // XmlHelper.dump(element);
      const rows = element.getElementsByTagName('a:tr');
      for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        // XmlHelper.dump(row);
        const columns = row.getElementsByTagName('a:tc');
        for (let c = 0; c < columns.length; c++) {
          const cell = columns.item(c);
          // XmlHelper.dump(cell);
          const texts = cell.getElementsByTagName('a:t');
          for (let t = 0; t < texts.length; t++) {
            data.push({
              row: r,
              column: c,
              text: texts.item(t).textContent,
            });
          }
        }
      }
    };

  await pres
    .addSlide('tables', 2, (slide) => {
      slide.modifyElement('NestedTable2', [readTableData(data)]);
      slide.modifyElement('NestedTable2', (element: XmlElement) => {
        data.forEach((tplCell) => {
          if (tplCell.text.indexOf('{{each') === 0) {
            if (tplCell.column > 0) {
              const tplCellXml = element
                .getElementsByTagName('a:tr')
                .item(tplCell.row)
                .getElementsByTagName('a:tc')
                .item(tplCell.column);

              vd(tplCell);
              XmlHelper.dump(tplCellXml);
            }
          }
        });
        // XmlHelper.dump(element);
      });
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  // console.log(data);

  // expect(data.length).toBe(12);
});
