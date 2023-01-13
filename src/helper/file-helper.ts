import fs from 'fs';
import path from 'path';
import JSZip, { InputType, JSZipObject, OutputType } from 'jszip';

import { AutomizerSummary, FileInfo } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { contentTracker } from './content-tracker';
import { CacheHelper } from './cache-helper';
import { vd } from './general-helper';
import { FileProxy } from './file-proxy';

export class FileHelper {
  static importArchive(location: string): FileProxy {
    if (!fs.existsSync(location)) {
      throw new Error('File not found: ' + location);
    }
    return FileProxy.importArchive(location);
  }

  static extractFromArchive(
    archive: FileProxy,
    file: string,
    type?: OutputType,
  ): Promise<FileProxy> {
    return archive.read(file, type);
  }

  static removeFromDirectory(
    archive: FileProxy,
    dir: string,
    cb: (file: JSZipObject, relativePath: string) => boolean,
  ): string[] {
    const removed = [];
    archive.folder(dir).forEach((relativePath, file) => {
      if (!relativePath.includes('/') && cb(file, relativePath)) {
        FileHelper.removeFromArchive(archive, file.name);
        removed.push(file.name);
      }
    });
    return removed;
  }

  static removeFromArchive(archive: FileProxy, file: string): FileProxy {
    FileHelper.check(archive, file);

    return archive.remove(file);
  }

  static getFileExtension(filename: string): string {
    return path.extname(filename).replace('.', '');
  }

  static getFileInfo(filename: string): FileInfo {
    return {
      base: path.basename(filename),
      dir: path.dirname(filename),
      isDir: filename[filename.length - 1] === '/',
      extension: path.extname(filename).replace('.', ''),
    };
  }

  static check(archive: FileProxy, file: string): boolean {
    FileHelper.isArchive(archive);
    return FileHelper.fileExistsInArchive(archive, file);
  }

  static isArchive(archive) {
    if (archive === undefined) {
      throw new Error('Archive is invalid or empty.');
    }
  }

  static fileExistsInArchive(archive: FileProxy, file: string): boolean {
    return archive.fileExists(file);
  }

  /**
   * Copies a file from one archive to another. The new file can have a different name to the origin.
   * @param {JSZip} sourceArchive - Source archive
   * @param {string} sourceFile - file path and name inside source archive
   * @param {JSZip} targetArchive - Target archive
   * @param {string} targetFile - file path and name inside target archive
   * @return {JSZip} targetArchive as an instance of JSZip
   */
  static async zipCopy(
    sourceArchive: FileProxy,
    sourceFile: string,
    targetArchive: FileProxy,
    targetFile?: string,
  ): Promise<FileProxy> {
    FileHelper.check(sourceArchive, sourceFile);

    const content = await sourceArchive.read(sourceFile, 'nodebuffer');
    contentTracker.trackFile(targetFile);

    return targetArchive.write(targetFile || sourceFile, content);
  }

  static async writeOutputFile(
    location: string,
    content: Buffer,
    automizer: IPresentationProps,
  ): Promise<AutomizerSummary> {
    await fs.promises.writeFile(location, content).catch((err) => {
      console.error(err);
      throw new Error(`Could not write output file: ${location}`);
    });

    const duration: number = (Date.now() - automizer.timer) / 600;

    return {
      status: 'finished',
      duration,
      file: location,
      filename: path.basename(location),
      templates: automizer.templates.length,
      slides: automizer.rootTemplate.count('slides'),
      charts: automizer.rootTemplate.count('charts'),
      images: automizer.rootTemplate.count('images'),
    };
  }
}

export const exists = (dir: string) => {
  return fs.existsSync(dir);
};

export const makeDirIfNotExists = (dir: string) => {
  if (!exists(dir)) {
    makeDir(dir);
  }
};

export const makeDir = (dir: string) => {
  try {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
  } catch (err) {
    throw err;
  }
};
