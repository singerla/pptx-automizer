import Automizer, { ConsoleLogger, ILogger, NullLogger } from '../src/index';

describe('ConsoleLogger verbosity filtering', () => {
  const spies = () => ({
    error: jest.spyOn(console, 'error').mockImplementation(),
    warn: jest.spyOn(console, 'warn').mockImplementation(),
    info: jest.spyOn(console, 'info').mockImplementation(),
    debug: jest.spyOn(console, 'debug').mockImplementation(),
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  test('verbosity 0 logs errors only', () => {
    const spy = spies();
    const logger = new ConsoleLogger(0);
    logger.error('e');
    logger.warn('w');
    logger.info('i');
    logger.debug('d');
    expect(spy.error).toHaveBeenCalledTimes(1);
    expect(spy.warn).not.toHaveBeenCalled();
    expect(spy.info).not.toHaveBeenCalled();
    expect(spy.debug).not.toHaveBeenCalled();
  });

  test('verbosity 1 adds warnings', () => {
    const spy = spies();
    const logger = new ConsoleLogger(1);
    logger.warn('w');
    logger.info('i');
    expect(spy.warn).toHaveBeenCalledTimes(1);
    expect(spy.info).not.toHaveBeenCalled();
  });

  test('verbosity 2 logs everything', () => {
    const spy = spies();
    const logger = new ConsoleLogger(2);
    logger.error('e');
    logger.warn('w');
    logger.info('i');
    logger.debug('d');
    expect(spy.error).toHaveBeenCalledTimes(1);
    expect(spy.warn).toHaveBeenCalledTimes(1);
    expect(spy.info).toHaveBeenCalledTimes(1);
    expect(spy.debug).toHaveBeenCalledTimes(1);
  });
});

test('an injected logger receives library output', async () => {
  const messages: { level: string; message: unknown }[] = [];
  const collect =
    (level: string) =>
    (message: unknown): void => {
      messages.push({ level, message });
    };
  const logger: ILogger = {
    error: collect('error'),
    warn: collect('warn'),
    info: collect('info'),
    debug: collect('debug'),
  };

  const result = await new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    logger,
    continueOnError: true,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('ThisShapeDoesNotExist', () => {});
    })
    .write(`logger-injection.test.pptx`);

  expect(result.slides).toBe(2);

  const warning = messages.find(
    (entry) =>
      entry.level === 'warn' &&
      String(entry.message).includes('ThisShapeDoesNotExist'),
  );
  expect(warning).toBeDefined();
});

test('NullLogger keeps the library silent', async () => {
  const spies = [
    jest.spyOn(console, 'error').mockImplementation(),
    jest.spyOn(console, 'warn').mockImplementation(),
    jest.spyOn(console, 'info').mockImplementation(),
    jest.spyOn(console, 'log').mockImplementation(),
    jest.spyOn(console, 'debug').mockImplementation(),
  ];

  await new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    logger: new NullLogger(),
    continueOnError: true,
  })
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('ThisShapeDoesNotExist', () => {});
    })
    .write(`logger-null.test.pptx`);

  spies.forEach((spy) => expect(spy).not.toHaveBeenCalled());
  jest.restoreAllMocks();
});
