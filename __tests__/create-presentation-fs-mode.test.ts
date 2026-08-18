import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import Automizer from '../src/automizer';

const outputDir = `${__dirname}/pptx-output`;

test('create presentation in fs mode and add basic slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir,
    archiveType: {
      mode: 'fs',
      baseDir: `${__dirname}/pptx-cache`,
      cleanupWorkDir: true,
    },
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes');

  pres.addSlide('shapes', 1);

  const outputFile = 'create-presentation-fs-mode.test.pptx';
  await pres.write(outputFile);

  const zip = await JSZip.loadAsync(
    fs.readFileSync(path.join(outputDir, outputFile)),
  );
  const slide = zip.file('ppt/slides/slide2.xml');
  expect(slide).not.toBeNull();
  expect(await slide!.async('text')).toContain('<p:sld');
});
