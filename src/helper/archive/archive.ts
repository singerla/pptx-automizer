import { DOMParser, Node, XMLSerializer } from '@xmldom/xmldom';
import { ArchivedFile, ArchiveType } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import {
  ArchiveParams,
  AutomizerFile,
  AutomizerParams,
} from '../../types/types';
import JSZip from 'jszip';

export default class Archive {
  filename: AutomizerFile;
  params: ArchiveParams;
  buffer: ArchivedFile[] = [];
  options: JSZip.JSZipGeneratorOptions<'nodebuffer'> = {
    type: 'nodebuffer',
  };

  constructor(filename: AutomizerFile, params: ArchiveParams) {
    this.filename = filename;
    this.params = params;
  }

  parseXml(xmlString: string): XmlDocument {
    const dom = new DOMParser();
    return dom.parseFromString(
      xmlString,
      'application/xml',
    ) as unknown as XmlDocument;
  }

  serializeXml(XmlDocument: XMLDocument | Node) {
    const s = new XMLSerializer();
    return s.serializeToString(<Node>XmlDocument);
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
