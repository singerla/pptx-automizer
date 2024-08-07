import Automizer, { modify, TableData } from './index';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
  });

  const tableData: TableData = {
    body: [
      { values: ['top left', 123, 345, 'subsub3-1', 'subsub3-2', 'Last'] },
      { values: [undefined, 't1', 't2', 't3', 't3', ''] },
      { values: ['label 0', 1, 2, 3, 3, 'l0'] },
      { values: ['label 1', 123, 345, 4563, 4671, 'l1'] },
      { values: ['label 2', 123, 345, 4562, 4672] },
      { values: ['label 3', 123, 345, 4561, 4673, 'l3'] },
      { values: [undefined, 'Foo', 'ter', 4564, 4674, ''] },
    ],
  };

  automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables')
    .addSlide('tables', 4, (slide) => {
      slide.modifyElement(
        'NestedTable3',
        modify.setTable(tableData, {
          adjustHeight: false,
          adjustWidth: false,
          expand: [
            {
              mode: 'row',
              tag: '{{each:row}}',
              count: 3,
            },
            {
              mode: 'column',
              tag: '{{each:subSub3}}',
              count: 1,
            },
          ],
        }),
      );
    })
    .write(`dev.pptx`)
    .then((summary) => {
      console.log(summary);
    });
};

run().catch((error) => {
  console.error(error);
});
