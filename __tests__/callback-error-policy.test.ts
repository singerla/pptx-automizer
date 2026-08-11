import Automizer, {
  AutomizerError,
  AutomizerParams,
  CallbackError,
  ElementNotFoundError,
} from '../src/index';

const setupPres = (params?: Partial<AutomizerParams>) =>
  new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    ...params,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes');

test('a throwing modification callback fails write() with CallbackError', async () => {
  const run = setupPres()
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('Cloud', () => {
        throw new Error('broken callback');
      });
    })
    .write(`callback-error-policy-loud.test.pptx`);

  await expect(run).rejects.toThrow(CallbackError);
  await expect(run).rejects.toThrow('broken callback');
  await expect(run).rejects.toThrow('Cloud');
});

test('CallbackError is an AutomizerError and preserves the cause', async () => {
  const cause = new Error('broken callback');
  const run = setupPres()
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('Cloud', () => {
        throw cause;
      });
    })
    .write(`callback-error-policy-cause.test.pptx`);

  const error = await run.catch((e) => e);
  expect(error).toBeInstanceOf(AutomizerError);
  expect(error).toBeInstanceOf(CallbackError);
  expect(error.cause).toBe(cause);
  expect(error.element).toBe('Cloud');
});

test('continueOnError skips a throwing modification callback', async () => {
  const result = await setupPres({ continueOnError: true, verbosity: 0 })
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('Cloud', () => {
        throw new Error('broken callback');
      });
    })
    .write(`callback-error-policy-lenient.test.pptx`);

  expect(result.slides).toBe(2);
});

test('an unresolvable element selector fails write() with ElementNotFoundError', async () => {
  const run = setupPres()
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('ThisShapeDoesNotExist', () => {});
    })
    .write(`callback-error-policy-selector-loud.test.pptx`);

  await expect(run).rejects.toThrow(ElementNotFoundError);
  await expect(run).rejects.toThrow('ThisShapeDoesNotExist');
});

test('continueOnError skips an unresolvable element selector', async () => {
  const result = await setupPres({ continueOnError: true, verbosity: 0 })
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('ThisShapeDoesNotExist', () => {});
    })
    .write(`callback-error-policy-selector-lenient.test.pptx`);

  expect(result.slides).toBe(2);
});
