import Automizer from '../src/automizer';
import ModifyPresentationHelper from '../src/helper/modify-presentation-helper';

test('read root presentation and sort slides', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`SlideWithImages.pptx`, 'images');

  for (let i = 0; i <= 2; i++) {
    pres.addSlide('shapes', 1);
  }

  pres.addSlide('charts', 1);
  pres.addSlide('charts', 2);
  pres.addSlide('images', 1);
  pres.addSlide('images', 2);

  const order = [2, 1, 8, 7, 5, 6, 3, 4];
  pres.modify(ModifyPresentationHelper.sortSlides(order));

  await pres.write(`sort-slides.test.pptx`);

  expect(pres).toBeInstanceOf(Automizer);
});
