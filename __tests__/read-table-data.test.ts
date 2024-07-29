import Automizer, { XmlElement } from '../src/index';
import { ChartModificationCallback } from '../src/types/types';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');
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
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableWithLabels', [readTableData(data)]);
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  // console.log(data);

  expect(data.length).toBe(12);
});
