import Archive from './archive';
import fs from 'fs';
import JSZip, { InputType } from 'jszip';
import { AutomizerParams } from '../../types/types';
import IArchive, { ArchivedFile } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';

export default class ArchiveJszip extends Archive implements IArchive {
  archive: JSZip;
  file: Buffer;
  options: JSZip.JSZipGeneratorOptions<'nodebuffer'> = {
    type: 'nodebuffer',
  };

  constructor(filename) {
    super(filename);
  }

  private async initialize() {
    this.file = await fs.promises.readFile(this.filename);

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

    const newArchive = new ArchiveJszip(file);
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

  async getContent(params: AutomizerParams): Promise<Buffer> {
    this.setOptions(params);

    await this.writeBuffer(this);

    return (await this.archive.generateAsync(this.options)) as Buffer;
  }

  private setOptions(params: AutomizerParams): void {
    if (params.compression > 0) {
      this.options.compression = 'DEFLATE';
      this.options.compressionOptions = {
        level: params.compression,
      };
    }
  }

  async readXml(file: string): Promise<XmlDocument> {
    const isBuffered = this.fromBuffer(file);

    if (!isBuffered) {
      const xmlString = (await this.read(file, 'string')) as string;
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
