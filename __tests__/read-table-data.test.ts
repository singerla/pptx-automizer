import Automizer from '../src/index';
import { TableInfo } from '../src/types/table-types';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  let data = <TableInfo[]>[];

  await pres
    .addSlide('tables', 1, async (slide) => {
      const tableInfo = await slide.getElement('TableWithLabels');

      data = tableInfo.getTableInfo();
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  console.log(data.map((data) => data.textContent));

  expect(data.length).toBe(12);
});
