import Automizer, { modify } from './index';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`EmptyTemplate.pptx`)
  .load(`SlideWithTables.pptx`, 'table');

const data1 = {
  body: [
    { label: 'item test r1', values: ['test1', 10, 16, 12] },
    { label: 'item test r2', values: ['test2', 12, 18, 15] },
    { label: 'item test r3', values: ['test3', 14, 12, 11] },
    // { label: 'item test r4', values: ['test4', 14, 12, 11] },
    // { label: 'item test r5', values: ['test5', 14, 12, 11] },
    // { label: 'item test r6', values: ['test6', 999, 12, 11] },
    // { label: 'item test r6', values: ['test7', 999, 12, 11] },
    // { label: 'item test r6', values: ['test8', 999, 12, 11] },
    // { label: 'item test r6', values: ['test9', 999, 12, 11] },
  ],
};

const data2 = [
  [10, 16, 12],
  [12, 18, 15],
  [14, 12, 11],
];

pres
  .addSlide('table', 1, (slide) => {
    slide.modifyElement('TableWithHeader', [modify.setTableData(data1)]);
  })
  .write(`modify-table.test.pptx`)

  .then((result) => {
    console.info(result);
  })
  .catch((error) => {
    console.error(error);
  });
