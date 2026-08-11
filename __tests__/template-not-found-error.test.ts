import Automizer, { TemplateNotFoundError } from '../src/index';

test('load() throws TemplateNotFoundError for a missing template file', () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  expect(() => automizer.load(`this-file-does-not-exist.pptx`)).toThrow(
    TemplateNotFoundError,
  );

  try {
    automizer.load(`this-file-does-not-exist.pptx`);
  } catch (error) {
    expect(error).toBeInstanceOf(TemplateNotFoundError);
    const templateError = error as TemplateNotFoundError;
    expect(templateError.file).toBe(`this-file-does-not-exist.pptx`);
    expect(templateError.message).toContain('pptx-templates');
  }
});

test('loadRoot() throws TemplateNotFoundError for a missing root template', () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  expect(() => automizer.loadRoot(`no-such-root.pptx`)).toThrow(
    TemplateNotFoundError,
  );
});
