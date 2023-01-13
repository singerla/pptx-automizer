import fs from 'fs';
import path from 'path';
import JSZip, { InputType, JSZipObject, OutputType } from 'jszip';

import { AutomizerSummary, FileInfo } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { contentTracker } from './content-tracker';
import { CacheHelper } from './cache-helper';
import { vd } from './general-helper';

export class FileProxy {
  archive: JSZip;
  file: any;
  filename: string;

  static importArchive(location: string): FileProxy {
    return new FileProxy(location);
  }

  constructor(filename) {
    this.filename = filename;
  }

  async inititalize() {
    this.file = await fs.promises.readFile(this.filename);

    const zip = new JSZip();
    this.archive = await zip.loadAsync(this.file as unknown as InputType);

    return this;
  }

  fileExists(file) {
    if (this.archive === undefined || this.archive.files[file] === undefined) {
      return false;
    }
    return true;
  }

  folder(dir) {
    return this.archive.folder(dir);
  }

  async count(pattern) {
    const files = (await this.filter(pattern)) as any;
    return files.length;
  }

  async read(file, type) {
    if (!this.archive) {
      await this.inititalize();
    }

    return this.archive.files[file].async(type || 'string');
  }

  async filter(pattern) {
    return this.archive.file(pattern);
  }

  async extract(file: string) {
    const contents = await this.read(file, 'nodebuffer');

    const zip = new JSZip();

    const newProxy = new FileProxy(file);
    newProxy.archive = await zip.loadAsync(contents as unknown as InputType);

    return newProxy;
  }

  async write(file, data) {
    this.archive.file(file, data);
    return this;
  }

  remove(file) {
    this.archive.remove(file);
    return this;
  }

  async send(options): Promise<Buffer> {
    return (await this.archive.generateAsync(options)) as Buffer;
  }
}
