import Automizer from './index';
import Slide from './slide';
import { dump, setPosition } from './helper/modify';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
});

const pres = automizer.loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithImages.pptx`, 'images')
  .load(`SlideWithLink.pptx`, 'link')
  .load(`SlideWithCharts.pptx`, 'charts')
  .load(`EmptySlide.pptx`, 'empty')
  .load(`SlideWithShapes.pptx`, 'shapes');

pres
  .addSlide('shapes', 2, (slide: Slide) => {
    slide.modifyElement('Drum', [dump, setPosition({x: 1000000, h: 5000000, w: 5000000})]);
  })

  .write(`myPresentation.pptx`).then(result => {
  console.info(result);
}).catch(error => {
  console.error(error);
});
