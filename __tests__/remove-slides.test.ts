import Automizer from '../src/automizer';
import ModifyPresentationHelper from '../src/helper/modify-presentation-helper';

test('Remove slides while and after modifying.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`SlideWithImages.pptx`, 'images');

  pres
    .addSlide('charts', 2, (slide) => {})
    .addSlide('images', 2, (slide) => {
      // Pass the current slide index to remove a slide.
      slide.remove(3);
      // You can as well remove another slide.
      slide.remove(1);
    })
    .addSlide('empty', 1, (slide) => {})
    .addSlide('empty', 1, (slide) => {})
    .addSlide('images', 1, (slide) => {
      // Calling
      slide.remove(1);
    });

  // Eventually remove some slides.
  pres.modify(ModifyPresentationHelper.removeSlides([1, 3]));

  const result = await pres.write(`remove-slides.test.pptx`);

  // ToDo: decrement counter
  expect(result.slides).toBe(6);
});
