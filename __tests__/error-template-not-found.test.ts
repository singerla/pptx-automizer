import Automizer from '../src/index';

test('throw error if template file not found', () => {
  const automizer = new Automizer({
    templateDir: `./pptx-templates`,
    outputDir: `./pptx-output`,
  });

  expect(() => {
    automizer.load(`non/existing/Template.pptx`);
  }).toThrow();
});
