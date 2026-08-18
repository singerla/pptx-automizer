import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import Automizer from '../src/automizer';

const outputDir = `${__dirname}/pptx-output`;

// fs-mode output is streamed to disk after write() resolves; poll until
// the file is a complete, parseable zip archive.
const waitForZip = async (file: string): Promise<JSZip> => {
  for (let i = 0; i < 150; i++) {
    if (fs.existsSync(file)) {
      try {
        return await JSZip.loadAsync(fs.readFileSync(file));
      } catch {
        // still being written
      }
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error('Output file was not written completely: ' + file);
};

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

  const zip = await waitForZip(path.join(outputDir, outputFile));
  const slide = zip.file('ppt/slides/slide2.xml');
  expect(slide).not.toBeNull();
  expect(await slide!.async('text')).toContain('<p:sld');
});
