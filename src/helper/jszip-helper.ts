import { log } from './logger';
import fs, { promises as fsp } from 'fs';
import { pipeline } from 'stream/promises';
import path from 'path';
import JSZip from 'jszip';

// Thanks to https://github.com/DesignByOnyx
// see https://github.com/Stuk/jszip/issues/386 for more info

/**
 * Returns a flat list of all files and subfolders for a directory (recursively).
 * @param {string} dir
 * @returns {Promise<string[]>}
 */
const getFilePathsRecursively = async (dir: string): Promise<string[]> => {
  // returns a flat array of absolute paths of all files recursively contained in the dir
  const list = await fsp.readdir(dir);
  const statPromises = list.map(async (file): Promise<string | string[]> => {
    const fullPath = path.resolve(dir, file);
    const stat = await fsp.stat(fullPath);
    if (stat && stat.isDirectory()) {
      return getFilePathsRecursively(fullPath);
    }
    return fullPath;
  });

  return (await Promise.all(statPromises)).flat() as string[];
};

/**
 * Creates an in-memory zip stream from a folder in the file system
 * @param {string} dir
 * @returns {JSZip}
 */
const createZipFromFolder = async (dir: string): Promise<JSZip> => {
  const absRoot = path.resolve(dir);
  const filePaths = await getFilePathsRecursively(dir);
  return filePaths.reduce((z, filePath) => {
    const relative = filePath.replace(absRoot, '');
    // create folder trees manually :(
    const zipFolder = path
      .dirname(relative)
      .split(path.sep)
      .reduce((zf, dirName) => zf.folder(dirName), z);

    zipFolder.file(path.basename(filePath), fs.createReadStream(filePath));
    return z;
  }, new JSZip());
};

const S_IFMT = 0xf000;
const S_IFLNK = 0xa000;

const isSymlinkEntry = (entry: JSZip.JSZipObject): boolean =>
  // unixPermissions is typed number | string | null; octal strings included
  (Number(entry.unixPermissions ?? 0) & S_IFMT) === S_IFLNK;

/**
 * Extracts a zip file into a folder in the file system.
 *
 * Every entry must stay inside destDir: absolute and path-traversal
 * ("zip slip") entry names are rejected, and symlink entries are skipped
 * (never created). The archive is held in memory while extracting, one
 * entry decompressed at a time - fine for .pptx-sized files.
 * @param {string} srcFile
 * @param {string} destDir
 */
export const extractToFolder = async (
  srcFile: string,
  destDir: string,
): Promise<void> => {
  const root = path.resolve(destDir);
  await fsp.mkdir(root, { recursive: true });

  const zip = await new JSZip().loadAsync(await fsp.readFile(srcFile));

  for (const entry of Object.values(zip.files)) {
    if (
      path.isAbsolute(entry.name) ||
      entry.name.startsWith('/') ||
      entry.name.startsWith('\\') ||
      /^[a-zA-Z]:/.test(entry.name)
    ) {
      throw new Error('Zip entry has an absolute path: ' + entry.name);
    }

    const dest = path.resolve(root, entry.name);
    if (dest !== root && !dest.startsWith(root + path.sep)) {
      throw new Error(
        'Zip entry resolves outside target directory: ' + entry.name,
      );
    }

    if (isSymlinkEntry(entry)) {
      log.warn('Skipping symlink entry in zip file: ' + entry.name);
      continue;
    }

    if (entry.dir) {
      await fsp.mkdir(dest, { recursive: true });
    } else {
      // some zips omit directory entries, so ensure the parent dir exists
      await fsp.mkdir(path.dirname(dest), { recursive: true });
      await fsp.writeFile(dest, await entry.async('nodebuffer'));
    }
  }
};

/**
 * Compresses a folder to the specified zip file.
 * @param {string} srcDir
 * @param {string} destFile
 */
export const compressFolder = async (
  srcDir: string,
  destFile: string,
  options: JSZip.JSZipGeneratorOptions<'nodebuffer'>,
) => {
  const start = Date.now();
  try {
    const zip = await createZipFromFolder(srcDir);
    // pipeline settles on completion or on an error from either stream, so
    // callers (e.g. ArchiveFs.output) only continue - and clean up the
    // source folder - once the zip is fully written
    await pipeline(
      zip.generateNodeStream({ streamFiles: true, ...options }),
      fs.createWriteStream(destFile),
    );
    log.info('Zip written successfully:', Date.now() - start, 'ms');
  } catch (ex) {
    log.error('Error creating zip', ex);
    throw ex;
  }
};
