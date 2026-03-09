import Automizer from '../src/automizer';

describe('Slide - Get Dimensions', () => {
  it('retrieves correct dimensions for slides', async () => {
    const automizer = new Automizer({
      templateDir: `${__dirname}/../__tests__/pptx-templates`,
      outputDir: `${__dirname}/../__tests__/pptx-output`,
      removeExistingSlides: true
    });

    const pres = automizer
      .loadRoot(`SlideDimensions1.pptx`)
      .load(`SlideDimensions1.pptx`, '1')
      .load(`SlideDimensions2.pptx`, '2');

    let dimensions1Promise = new Promise((resolve, reject) => {
      pres.addSlide('1', 1, async (slide) => {
        try {
          const dimensions = await slide.getDimensions();
          resolve(dimensions);
        } catch (error) {
          reject(error);
        }
      });
    });

    let dimensions2Promise = new Promise((resolve, reject) => {
      pres.addSlide('2', 1, async (slide) => {
        try {
          const dimensions = await slide.getDimensions();
          resolve(dimensions);
        } catch (error) {
          reject(error);
        }
      });
    });

    const dimensions1 = await dimensions1Promise;
    const dimensions2 = await dimensions2Promise;

    expect(dimensions2).toEqual({ width: 9323387, height: 5670550 });
    expect(dimensions1).toEqual({ width: 10080625, height: 6300787 });

    await pres.write(`slideGetDimensionsTestOutput.pptx`);
  });
});
