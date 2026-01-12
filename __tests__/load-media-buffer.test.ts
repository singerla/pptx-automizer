import Automizer from '../src/automizer';
import { ModifyImageHelper } from '../src';
import * as fs from 'fs';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const mediaDir = `${__dirname}/media`;

describe('loadMediaBuffer', () => {
  test('load single media file from buffer', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);
    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer('logo.png', buffer);

    expect(automizer.rootTemplate.mediaFiles.length).toBe(1);
    expect(automizer.rootTemplate.mediaFiles[0].source).toBe('buffer');
    expect(automizer.rootTemplate.mediaFiles[0].file).toBe('logo.png');
    expect(automizer.rootTemplate.mediaFiles[0].extension).toBe('png');
  });

  test('load multiple media files from buffers', async () => {
    const buffer1 = fs.readFileSync(`${mediaDir}/test.png`);
    const buffer2 = fs.readFileSync(`${mediaDir}/feather.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer(['logo.png', 'icon.png'], [buffer1, buffer2]);

    expect(automizer.rootTemplate.mediaFiles.length).toBe(2);
    expect(automizer.rootTemplate.mediaFiles[0].source).toBe('buffer');
    expect(automizer.rootTemplate.mediaFiles[0].file).toBe('logo.png');
    expect(automizer.rootTemplate.mediaFiles[1].source).toBe('buffer');
    expect(automizer.rootTemplate.mediaFiles[1].file).toBe('icon.png');
  });

  test('load buffer with prefix', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer('logo.png', buffer, 'test_');

    expect(automizer.rootTemplate.mediaFiles.length).toBe(1);
    expect(automizer.rootTemplate.mediaFiles[0].prefix).toBe('test_');
  });

  test('load both file-based and buffer-based media', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
      mediaDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMedia('feather.png')
      .loadMediaBuffer('buffer.png', buffer);

    expect(automizer.rootTemplate.mediaFiles.length).toBe(2);
    expect(automizer.rootTemplate.mediaFiles[0].source).toBe('path');
    expect(automizer.rootTemplate.mediaFiles[0].file).toBe('feather.png');
    expect(automizer.rootTemplate.mediaFiles[1].source).toBe('buffer');
    expect(automizer.rootTemplate.mediaFiles[1].file).toBe('buffer.png');
  });

  test('throw error for filename without extension', () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    }).loadRoot('RootTemplate.pptx');

    expect(() => {
      automizer.loadMediaBuffer('logo', buffer);
    }).toThrow('Filename must include extension');
  });

  test('throw error for empty buffer', () => {
    const automizer = new Automizer({
      templateDir,
      outputDir,
    }).loadRoot('RootTemplate.pptx');

    expect(() => {
      automizer.loadMediaBuffer('logo.png', Buffer.alloc(0));
    }).toThrow('Empty buffer provided');
  });

  test('throw error for invalid buffer', () => {
    const automizer = new Automizer({
      templateDir,
      outputDir,
    }).loadRoot('RootTemplate.pptx');

    expect(() => {
      automizer.loadMediaBuffer('logo.png', 'not a buffer' as any);
    }).toThrow('Invalid buffer for file');
  });

  test('throw error for duplicate filename', () => {
    const buffer1 = fs.readFileSync(`${mediaDir}/test.png`);
    const buffer2 = fs.readFileSync(`${mediaDir}/feather.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    }).loadRoot('RootTemplate.pptx');

    automizer.loadMediaBuffer('logo.png', buffer1);

    expect(() => {
      automizer.loadMediaBuffer('logo.png', buffer2);
    }).toThrow('already loaded');
  });

  test('throw error for mismatched arrays', () => {
    const buffer1 = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    }).loadRoot('RootTemplate.pptx');

    expect(() => {
      automizer.loadMediaBuffer(['a.png', 'b.png'], [buffer1]);
    }).toThrow('Mismatched arrays');
  });

  test('throw error when root template not loaded', () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    });

    expect(() => {
      automizer.loadMediaBuffer('logo.png', buffer);
    }).toThrow("Can't load media, you need to load a root template first");
  });

  test('reference buffer-loaded image with setRelationTarget', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer('custom.png', buffer)
      .load('SlideWithImages.pptx', 'images');

    const pres = automizer.addSlide('images', 1, (slide) => {
      slide.modifyElement('Grafik 5', [
        ModifyImageHelper.setRelationTarget('custom.png'),
      ]);
    });

    const result = await pres.write(`buffer-image-test.pptx`);
    expect(result.images).toBeGreaterThanOrEqual(1);
  });

  test('use setRelationTargetCover with buffer-loaded media', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer('new-image.png', buffer)
      .load('SlideWithImages.pptx', 'images');

    const pres = automizer.addSlide('images', 1, (slide) => {
      slide.modifyElement('Grafik 5', [
        ModifyImageHelper.setRelationTargetCover('new-image.png', automizer),
      ]);
    });

    const result = await pres.write(`buffer-image-cover-test.pptx`);
    expect(result.images).toBeGreaterThanOrEqual(1);
  });

  test('mixed file and buffer loading with different prefixes', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
      mediaDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMedia('feather.png', undefined, 'file_')
      .loadMediaBuffer('test.png', buffer, 'buffer_');

    expect(automizer.rootTemplate.mediaFiles.length).toBe(2);
    expect(automizer.rootTemplate.mediaFiles[0].prefix).toBe('file_');
    expect(automizer.rootTemplate.mediaFiles[1].prefix).toBe('buffer_');
  });

  test('write buffer-loaded media to archive', async () => {
    const buffer = fs.readFileSync(`${mediaDir}/test.png`);

    const automizer = new Automizer({
      templateDir,
      outputDir,
    })
      .loadRoot('RootTemplate.pptx')
      .loadMediaBuffer('myimage.png', buffer)
      .load('SlideWithImages.pptx', 'images');

    const pres = automizer.addSlide('images', 1, (slide) => {
      slide.modifyElement('Grafik 5', [
        ModifyImageHelper.setRelationTarget('myimage.png'),
      ]);
    });

    const result = await pres.write(`buffer-write-test.pptx`);
    expect(result.images).toBeGreaterThanOrEqual(1);

    // Verify the output file exists
    const outputPath = `${outputDir}/buffer-write-test.pptx`;
    expect(fs.existsSync(outputPath)).toBe(true);
  });
});
