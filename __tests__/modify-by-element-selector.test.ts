import Automizer, { modify } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('create presentation, add and modify an existing table by FindElementSelector', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    useCreationIds: true,
  });

  const data1 = {
    body: [
      { label: 'item test r1', values: ['test1', 10, 16, 12, 11] },
      { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
      { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTables.pptx`, 'tables');

  const creationIds = await pres.setCreationIds();
  // vd(creationIds);

  const result = await pres
    .addSlide('tables', 1950777067, (slide) => {
      // This will try to match the given creationId first, and, if failed,
      // match by element name.
      slide.modifyElement(
        {
          creationId: '{EFC74B4C-D832-409B-9CF4-73C1EFF132D8}',
          name: 'TableDefault',
        },
        [modify.setTableData(data1)],
      );

      // CreationID is not valid/has changed, but we can use a fallback element name:
      slide.addElement(
        'tables',
        1950777067,
        {
          creationId: '{XXXXX-XXXXX-XXXXX-XX-XXX}',
          name: 'TableWithLabels',
        },
        [modify.setTableData(data1)],
      );
    })
    .write(`modify-by-element-selector.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
