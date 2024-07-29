import * as fs from 'fs/promises';
import Automizer from '../src/automizer';

test('create presentation from buffer and add basic slide', async () => {
  const automizer = new Automizer({
    outputDir: `${__dirname}/pptx-output`,
  });
  const rootTemplate = await fs.readFile(
    `${__dirname}/pptx-templates/RootTemplate.pptx`,
  );
  const slideWithShapes = await fs.readFile(
    `${__dirname}/pptx-templates/SlideWithShapes.pptx`,
  );
  const pres = automizer.loadRoot(rootTemplate).load(slideWithShapes, 'shapes');

  for (let i = 0; i <= 10; i++) {
    pres.addSlide('shapes', 1);
  }

  await pres.write(`create-presentation-from-buffer.test.pptx`);

  expect(pres).toBeInstanceOf(Automizer);
});
