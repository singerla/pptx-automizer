import fs from 'fs';
import { promises as fsPromises } from 'fs';
import path from 'path';
import { FileInfo, ArchiveParams, AutomizerFile } from '../types/types';
import IArchive, {
  ArchivedFolderCallback,
} from '../interfaces/iarchive';
import ArchiveJszip from './archive/archive-jszip';
import ArchiveFs from './archive/archive-fs';
import { ContentTypeExtension } from '../enums/content-type-map';

/**
 * The slice of Slide/Master/Layout/Shape that archive copy helpers need:
 * the two archives a part is copied between.
 */
export type ArchiveCopyContext = {
  sourceArchive: IArchive;
  targetArchive: IArchive;
};

export class FileHelper {
  static importArchive(file: AutomizerFile, params: ArchiveParams): IArchive {
    if (typeof file !== 'object') {
      if (!fs.existsSync(file)) {
        throw new Error('File not found: ' + file);
      }

      switch (params.mode) {
        case 'jszip':
          return new ArchiveJszip(file, params);
        case 'fs':
          return new ArchiveFs(file, params);
      }
    } else {
      return new ArchiveJszip(file, params);
    }
  }

  static async removeFromDirectory(
    archive: IArchive,
    dir: string,
    cb: ArchivedFolderCallback,
  ): Promise<string[]> {
    const removed = [];
    const files = await archive.folder(dir);
    for (const file of files) {
      if (cb(file)) {
        await archive.remove(file.name);
        removed.push(file.name);
      }
    }

    return removed;
  }

  static getFileExtension(filename: string): ContentTypeExtension {
    return path.extname(filename).replace('.', '') as ContentTypeExtension;
  }

  static getFileInfo(filename: string): FileInfo {
    return {
      base: path.basename(filename),
      dir: path.dirname(filename),
      isDir: filename[filename.length - 1] === '/',
      extension: path.extname(filename).replace('.', ''),
    };
  }

  static check(archive: IArchive, file: string): boolean {
    FileHelper.isArchive(archive);
    return FileHelper.fileExistsInArchive(archive, file);
  }

  static isArchive(archive: IArchive): void {
    if (archive === undefined) {
      throw new Error('Archive is invalid or empty.');
    }
  }

  static fileExistsInArchive(archive: IArchive, file: string): boolean {
    return archive.fileExists(file);
  }

  static async zipCopyWithRelations(
    context: ArchiveCopyContext,
    type: string,
    sourceNumber: number,
    targetNumber: number,
  ): Promise<void> {
    const typePlural = type + 's';
    await FileHelper.zipCopyByIndex(
      context,
      `ppt/${typePlural}/${type}`,
      sourceNumber,
      targetNumber,
    );
    await FileHelper.zipCopyByIndex(
      context,
      `ppt/${typePlural}/_rels/${type}`,
      sourceNumber,
      targetNumber,
      '.xml.rels',
    );
  }

  static async zipCopyByIndex(
    context: ArchiveCopyContext,
    prefix: string,
    sourceId: number | string,
    targetId: number | string,
    suffix?: string,
  ): Promise<IArchive> {
    suffix = suffix || '.xml';
    return FileHelper.zipCopy(
      context.sourceArchive,
      `${prefix}${sourceId}${suffix}`,
      context.targetArchive,
      `${prefix}${targetId}${suffix}`,
    );
  }

  /**
   * Copies a file from one archive to another. The new file can have a different name to the origin.
   * @param {IArchive} sourceArchive - Source archive
   * @param {string} sourceFile - file path and name inside source archive
   * @param {IArchive} targetArchive - Target archive
   * @param {string} targetFile - file path and name inside target archive
   * @return {IArchive} targetArchive as an instance of IArchive
   */
  static async zipCopy(
    sourceArchive: IArchive,
    sourceFile: string,
    targetArchive: IArchive,
    targetFile?: string,
  ): Promise<IArchive> {
    FileHelper.check(sourceArchive, sourceFile);
    targetArchive.contentTracker?.trackFile(targetFile);

    const content = await sourceArchive
      .read(sourceFile, 'nodebuffer')
      .catch((e) => {
        throw e;
      });

    return targetArchive.write(targetFile || sourceFile, content);
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
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir);
  }
};

export const copyDir = async (src: string, dest: string) => {
  await fsPromises.mkdir(dest, { recursive: true });
  const entries = await fsPromises.readdir(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      await copyDir(srcPath, destPath);
    } else {
      await fsPromises.copyFile(srcPath, destPath);
    }
  }
};

export const ensureDirectoryExistence = (filePath: string): void => {
  const dirname = path.dirname(filePath);
  if (fs.existsSync(dirname)) {
    return;
  }
  ensureDirectoryExistence(dirname);
  fs.mkdirSync(dirname);
};
