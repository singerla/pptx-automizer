import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { compressFolder, extractToFolder } from '../src/helper/jszip-helper';

const cacheDir = `${__dirname}/pptx-cache/extract-to-folder`;

const writeZip = async (
  zip: JSZip,
  name: string,
  options: Partial<JSZip.JSZipGeneratorOptions<'nodebuffer'>> = {},
): Promise<string> => {
  fs.mkdirSync(cacheDir, { recursive: true });
  const file = path.join(cacheDir, name);
  fs.writeFileSync(
    file,
    await zip.generateAsync({ type: 'nodebuffer', ...options }),
  );
  return file;
};

afterAll(() => {
  fs.rmSync(cacheDir, { recursive: true, force: true });
});

describe('extractToFolder', () => {
  test('extracts nested files, including entries without explicit directories', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.folder('docProps').file('core.xml', '<coreProperties/>');
    // no explicit directory entry for ppt/slides
    zip.file('ppt/slides/slide1.xml', '<p:sld/>');

    const srcFile = await writeZip(zip, 'roundtrip.zip');
    const target = path.join(cacheDir, 'roundtrip');
    await extractToFolder(srcFile, target);

    expect(
      fs.readFileSync(path.join(target, '[Content_Types].xml'), 'utf8'),
    ).toBe('<Types/>');
    expect(
      fs.readFileSync(path.join(target, 'docProps/core.xml'), 'utf8'),
    ).toBe('<coreProperties/>');
    expect(
      fs.readFileSync(path.join(target, 'ppt/slides/slide1.xml'), 'utf8'),
    ).toBe('<p:sld/>');
  });

  // JSZip normalizes traversal names on creation, so the malicious archives
  // below are forged by patching the entry name bytes after generation
  // (same name length keeps all zip offsets valid).
  const forgeZip = async (
    originalName: string,
    forgedName: string,
    zipName: string,
  ): Promise<string> => {
    expect(forgedName.length).toBe(originalName.length);
    const zip = new JSZip();
    zip.file(originalName, 'pwned', { createFolders: false });
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const patched = Buffer.from(
      buffer.toString('latin1').split(originalName).join(forgedName),
      'latin1',
    );
    fs.mkdirSync(cacheDir, { recursive: true });
    const srcFile = path.join(cacheDir, zipName);
    fs.writeFileSync(srcFile, patched);
    return srcFile;
  };

  test('never writes path traversal ("zip slip") entries outside the target', async () => {
    const srcFile = await forgeZip('AA/evil.txt', '../evil.txt', 'zip-slip.zip');
    const target = path.join(cacheDir, 'zip-slip');

    // jszip >= 3.8 already strips ".." on load ("../evil.txt" -> "evil.txt");
    // the extractToFolder guard is defense in depth on top of that. Either
    // way, nothing may be written outside the target directory.
    await extractToFolder(srcFile, target).catch((err) => {
      expect(String(err)).toContain('outside target directory');
    });
    expect(fs.existsSync(path.join(cacheDir, 'evil.txt'))).toBe(false);
    expect(fs.existsSync(path.join(cacheDir, '..', 'evil.txt'))).toBe(false);
  });

  test('rejects entries with a leading slash (kept as-is by jszip)', async () => {
    const srcFile = await forgeZip('Xabs.txt', '/abs.txt', 'leading-slash.zip');

    await expect(
      extractToFolder(srcFile, path.join(cacheDir, 'leading-slash')),
    ).rejects.toThrow('Zip entry has an absolute path: /abs.txt');
    expect(fs.existsSync('/abs.txt')).toBe(false);
  });

  test('rejects absolute entry paths', async () => {
    const zip = new JSZip();
    zip.file('C:\\abs.txt', 'x');

    const srcFile = await writeZip(zip, 'absolute.zip');
    await expect(
      extractToFolder(srcFile, path.join(cacheDir, 'absolute')),
    ).rejects.toThrow('Zip entry has an absolute path: C:\\abs.txt');
  });

  test('compressFolder rejects when the source folder cannot be read', async () => {
    // a failed write must reach the caller (ArchiveFs.output) instead of
    // reporting success and cleaning up over a truncated output file
    await expect(
      compressFolder(
        path.join(cacheDir, 'does-not-exist'),
        path.join(cacheDir, 'never-written.zip'),
        {},
      ),
    ).rejects.toThrow();
  });

  test('skips symlink entries instead of creating them', async () => {
    const zip = new JSZip();
    zip.file('regular.txt', 'content');
    zip.file('link.txt', '../../outside-target', {
      unixPermissions: 0o120777, // S_IFLNK
    });

    const srcFile = await writeZip(zip, 'symlink.zip', { platform: 'UNIX' });
    const target = path.join(cacheDir, 'symlink');
    await extractToFolder(srcFile, target);

    expect(fs.readFileSync(path.join(target, 'regular.txt'), 'utf8')).toBe(
      'content',
    );
    // neither a symlink nor a regular file may be created for the entry
    expect(fs.existsSync(path.join(target, 'link.txt'))).toBe(false);
    expect(() => fs.lstatSync(path.join(target, 'link.txt'))).toThrow();
  });
});
