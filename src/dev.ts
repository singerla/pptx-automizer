import Automizer, { modify } from './index';
import { vd } from './helper/general-helper';
import * as fs from 'fs';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  // Step 1: Create a pptx with images and a chart inside.
  // The chart is modified by pptx-automizer

  const writer = new Automizer({
    templateDir,
    outputDir,
    verbosity: 0,
    removeExistingSlides: true
  });

  const pres = writer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithImages.pptx`, 'images')
    .load(`ChartBarsStacked.pptx`, 'charts');

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, null] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('images', 1, 'Grafik 5');
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .write(`modify-automizer-generated-file.tmp.test.pptx`);

  // Step 2: Create a copy of the generated file in templateDir
  // and load it as a normal template

  await fs.promises.copyFile(
    `${outputDir}/modify-automizer-generated-file.tmp.test.pptx`,
    `${templateDir}/PptxAutomizerGeneratedFile.pptx`,
  );

  const reader = new Automizer({
    templateDir,
    outputDir,
    // This will display all log() output
    verbosity: 2
  });


  const dataSmallerMod = {
    series: [{ label: 'series s3' }, { label: 'series s4' }],
    categories: [
      { label: 'item test r3', values: [22, 45] },
      { label: 'item test r4', values: [23, 46] },
      { label: 'item test r5', values: [24, 47] },
    ],
  };

  const pres2 = reader
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithImages.pptx`, 'images')
    .load(`PptxAutomizerGeneratedFile.pptx`, 'generated')

  const presInfo = await pres2.getInfo();
  const slides = presInfo
    .slidesByTemplate(`generated`)

  console.log(`The re-imported file PptxAutomizerGeneratedFile.pptx seems to have ${slides.length} slides`)
  // But only 2 slides were expected with removeExistingSlides: true

  const result2 = await pres2
    .addSlide('empty', 1, (slide) => {
      slide.addElement('images', 1, 'Grafik 5');
    })
    .addSlide('generated', 2, (slide) => {
      slide.addElement('images', 1, 'Grafik 5');
    })
    .addSlide('generated', 3, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmallerMod)]);
    })
    .write(`modify-automizer-generated-file.test.pptx`);

  vd(result2)

};

run().catch((error) => {
  console.error(error);
});
