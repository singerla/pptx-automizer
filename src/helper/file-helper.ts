import fs from 'fs';
import path from 'path';
import JSZip, { InputType, OutputType } from 'jszip';

import { AutomizerSummary } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';

export class FileHelper {
  static readFile(location: string): Promise<Buffer> {
    if (!fs.existsSync(location)) {
      throw new Error('File not found: ' + location);
    }
    return fs.promises.readFile(location);
  }

  static extractFromArchive(
    archive: JSZip,
    file: string,
    type?: OutputType,
  ): Promise<string | number[] | Uint8Array | ArrayBuffer | Blob | Buffer> {
    if (archive === undefined) {
      throw new Error('No files found, expected: ' + file);
    }

    if (archive.files[file] === undefined) {
      console.trace();
      throw new Error('Archived file not found: ' + file);
    }
    return archive.files[file].async(type || 'string');
  }

  static extractFileContent(file: Buffer): Promise<JSZip> {
    const zip = new JSZip();
    return zip.loadAsync((file as unknown) as InputType);
  }

  static getFileExtension(filename: string): string {
    return path.extname(filename).replace('.', '');
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
    sourceArchive: JSZip,
    sourceFile: string,
    targetArchive: JSZip,
    targetFile?: string,
  ): Promise<JSZip> {
    if (sourceArchive.files[sourceFile] === undefined) {
      throw new Error(`Zipped file not found: ${sourceFile}`);
    }

    const content = sourceArchive.files[sourceFile].async('nodebuffer');
    return targetArchive.file(targetFile || sourceFile, content);
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
