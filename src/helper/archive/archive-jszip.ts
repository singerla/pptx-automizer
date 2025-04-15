import Archive from './archive';
import fs from 'fs';
import JSZip, { InputType } from 'jszip';
import {
  ArchiveParams,
  AutomizerFile,
  AutomizerParams,
} from '../../types/types';
import IArchive, { ArchivedFile } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import path from 'path';
import { vd } from '../general-helper';

export default class ArchiveJszip extends Archive implements IArchive {
  archive: JSZip;
  file: Buffer;

  constructor(filename: AutomizerFile, params: ArchiveParams) {
    super(filename, params);
  }

  private async initialize() {
    if (typeof this.filename !== 'object') {
      this.file = await fs.promises.readFile(this.filename);
    } else {
      this.file = this.filename as Buffer;
    }
    const zip = new JSZip();
    this.archive = await zip.loadAsync(this.file as unknown as InputType);

    return this;
  }

  fileExists(file: string) {
    if (this.archive === undefined || this.archive.files[file] === undefined) {
      return false;
    }
    return true;
  }

  async folder(dir: string): Promise<ArchivedFile[]> {
    const files = <ArchivedFile[]>[];
    this.archive.folder(dir).forEach((relativePath, file) => {
      if (!relativePath.includes('/')) {
        files.push({
          name: file.name,
          relativePath,
        });
      }
    });
    return files;
  }

  async read(
    file: string,
    type: 'string' | 'nodebuffer',
  ): Promise<string | Buffer> {
    if (!this.archive) {
      await this.initialize();
    }

    if (!this.archive.files[file]) {
      if (typeof this.filename === 'string') {
        throw new Error(
          'Could not find file ' + file + '@' + path.basename(this.filename),
        );
      } else {
        throw new Error('Could not find file ' + file);
      }
    }

    return this.archive.files[file].async(type || 'string');
  }

  async write(file: string, data: string | Buffer): Promise<this> {
    this.archive.file(file, data);
    return this;
  }

  async remove(file: string): Promise<void> {
    this.archive.remove(file);
  }

  async extract(file: string): Promise<ArchiveJszip> {
    const contents = (await this.read(file, 'nodebuffer')) as Buffer;

    const zip = new JSZip();

    const newArchive = new ArchiveJszip(file, this.params);
    newArchive.archive = await zip.loadAsync(contents as unknown as InputType);

    return newArchive;
  }

  async output(location: string, params: AutomizerParams): Promise<void> {
    const content = await this.getContent(params);

    await fs.promises.writeFile(location, content).catch((err) => {
      console.error(err);
      throw new Error(`Could not write output file: ${location}`);
    });
  }

  async stream(
    params: AutomizerParams,
    options?: JSZip.JSZipGeneratorOptions<'nodebuffer'>,
  ): Promise<NodeJS.ReadableStream> {
    this.setOptions(params);

    await this.writeBuffer(this);

    const mergedOptions = {
      ...this.options,
      ...options,
    };

    return this.archive.generateNodeStream(mergedOptions);
  }

  async getFinalArchive(): Promise<JSZip> {
    await this.writeBuffer(this);
    return this.archive;
  }

  async getContent(params: AutomizerParams): Promise<Buffer> {
    this.setOptions(params);

    await this.writeBuffer(this);

    return (await this.archive.generateAsync(this.options)) as Buffer;
  }

  async readXml(file: string): Promise<XmlDocument> {
    const isBuffered = this.fromBuffer(file);

    if (!isBuffered) {
      let xmlString: string = '';
      if (this.params.decodeText) {
        const buffer = (await this.read(file, 'nodebuffer')) as Buffer;
        xmlString = new TextDecoder().decode(buffer);
      } else {
        xmlString = (await this.read(file, 'string')) as string;
      }

      const XmlDocument = this.parseXml(xmlString);
      this.toBuffer(file, XmlDocument);

      return XmlDocument;
    } else {
      return isBuffered.content;
    }
  }

  writeXml(file: string, XmlDocument: XmlDocument): void {
    this.toBuffer(file, XmlDocument);
  }
}
