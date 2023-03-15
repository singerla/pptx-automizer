import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { ArchivedFile, ArchiveType } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import { AutomizerParams } from '../../types/types';
import JSZip from 'jszip';

export default class Archive {
  filename: string;
  buffer: ArchivedFile[] = [];
  options: JSZip.JSZipGeneratorOptions<'nodebuffer'> = {
    type: 'nodebuffer',
  };

  constructor(filename) {
    this.filename = filename;
  }

  parseXml(xmlString: string): XmlDocument {
    const dom = new DOMParser();
    return dom.parseFromString(xmlString);
  }

  serializeXml(XmlDocument: XmlDocument) {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(XmlDocument);
    return xmlBuffer;
  }

  async writeBuffer(archiveType: ArchiveType) {
    for (const buffered of this.buffer) {
      const serialized = this.serializeXml(buffered.content);
      await archiveType.write(buffered.relativePath, serialized);
    }
  }

  toBuffer(relativePath, content): void {
    const existing = this.fromBuffer(relativePath);
    if (!existing) {
      this.buffer.push({
        relativePath,
        name: relativePath,
        content: content,
      });
    }
  }

  setOptions(params: AutomizerParams): void {
    if (params.compression > 0) {
      this.options.compression = 'DEFLATE';
      this.options.compressionOptions = {
        level: params.compression,
      };
    }
  }

  fromBuffer(relativePath) {
    return this.buffer.find((file) => file.relativePath === relativePath);
  }
}
