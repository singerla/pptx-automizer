import Automizer, { TableData } from '../src/index';
import { ModifyTableHelper } from '../src';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const data: TableData = {
    body: [],
  };

  await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableWithLabels', [
        ModifyTableHelper.readTableData(data),
      ]);
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  // console.log(data);
  expect(data.body.length).toBe(12);
});
