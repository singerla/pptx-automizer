import Automizer, { TableData } from '../src/index';
import { ModifyTableHelper } from '../src';
import { TableInfo } from '../src/types/table-types';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const data = <TableInfo[]>[];

  await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableWithLabels', [
        // ToDo: use from ElementInfo
        // ModifyTableHelper.readTableData(data),
      ]);
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  // console.log(data);
  expect(data.length).toBe(12);
});
