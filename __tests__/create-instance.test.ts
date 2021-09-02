import Automizer from '../src/index';

test('create automizer instance', () => {
  const automizer = new Automizer({
    templateDir: `./pptx-templates`,
    outputDir: `./pptx-output`,
  });

  expect(automizer).toBeInstanceOf(Automizer);
});
