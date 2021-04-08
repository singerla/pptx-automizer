import Automizer, { modify } from './index';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithTables.pptx`, 'table');

// const data1 = {
//   body: [
//     { label: 'item test r1', values: [10, 16, 12] },
//     { label: 'item test r2', values: [12, 18, 15] },
//     { label: 'item test r3', values: [14, 12, 11] },
//   ],
// };

const data2 = [
  [10, 16, 12],
  [12, 18, 15],
  [14, 12, 11],
];

pres
  .addSlide('table', 1, (slide) => {
    slide.modifyElement('TableWithHeader', [modify.setTableData(data2)]);
  })
  .write(`modify-table.test.pptx`)

  .then((result) => {
    console.info(result);
  })
  .catch((error) => {
    console.error(error);
  });
