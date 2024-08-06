import Automizer, { XmlElement } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('read table data from slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables');

  await pres
    .addSlide('tables', 2, async (slide) => {
      const info = await slide.getElement('NestedTable2');
      const data = info.getTableData().body;

      slide.modifyElement('NestedTable2', (element: XmlElement) => {
        data.forEach((tplCell) => {
          vd(tplCell);
          // Test
        });
        // XmlHelper.dump(element);
      });
    })
    .write(`read-table-data.test.pptx`);

  // We have 12 text values in a 3x3 table:
  // console.log(data);

  // expect(data.length).toBe(12);
});
