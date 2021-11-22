import Automizer, { modify } from './index';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`EmptyTemplate.pptx`)
  .load(`TextReplace.pptx`)
  // .load(`Library - Retailer.pptx`)
  // .load(`SlideWithTables.pptx`, 'table')
  // .load('SlideWithImages.pptx')
  // .load(`RootTemplateWithCharts.pptx`);


//
// pres.getCreationIds()
//   .then(response => {
//     console.dir(response, {depth: 5})
//   })

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

const run = async() => {
  // const info = await pres.setCreationIds()
  // console.dir(info, {depth: 5})

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'Text',
        modify.replaceText([
          {
            replace: 'replace',
            by: {
              text: 'Apples'
            }
          },
          {
            replace: 'by',
            by: {
              text: 'Bananas'
            }
          },
          {
            replace: 'replacement',
            by: [
              {
                text: 'Really!',
                style: {
                  size: 10000
                }
              },
              {
                text: 'Fine!',
                style: {
                  size: 10000
                }
              }
            ]
          },
        ], {
          openingTag: '${',
          closingTag: '}'
        })
      )
    })
    .write(`text-replacement.test.pptx`)
    .then((result) => {
      // console.info(result);
    })

  return pres
}

run()
  .catch((error) => {
    console.error(error);
  });

