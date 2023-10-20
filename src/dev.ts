import Automizer, {
  CmToDxa,
  ISlide,
  ModifyColorHelper,
  ModifyShapeHelper,
  ModifyTextHelper,
} from './index';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`SlideWithShapes.pptx`)
    // We load it twice to make it available for modifying slides
    // Defining a "name" as second params makes it a little easier
    .load(`SlideWithShapes.pptx`, 'myTemplate');

  // This is brandnew: get useful information about loaded templates:
  const myTemplates = await pres.getInfo();
  const mySlides = myTemplates.slidesByTemplate(`myTemplate`);

  // Feel free to create some functions to pre-define all modifications
  // you need to apply to your slides.
  type CallbackBySlideNumber = {
    slideNumber: number;
    callback: (slide: ISlide) => void;
  };
  const callbacks: CallbackBySlideNumber[] = [
    {
      slideNumber: 2,
      callback: (slide: ISlide) => {
        slide.modifyElement('Cloud', [
          ModifyTextHelper.setText('My content'),
          ModifyShapeHelper.setPosition({
            h: CmToDxa(5),
          }),
          ModifyColorHelper.solidFill({
            type: 'srgbClr',
            value: 'cccccc',
          }),
        ]);
      },
    },
  ];
  const getCallbacks = (slideNumber: number) => {
    return callbacks.find((callback) => callback.slideNumber === slideNumber)
      ?.callback;
  };

  // We can loop all slides an apply the callbacks if defined
  mySlides.forEach((mySlide) => {
    pres.addSlide('myTemplate', mySlide.number, getCallbacks(mySlide.number));
  });

  // This will result to an output presentation containing all slides of "SlideWithShapes.pptx"
  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
