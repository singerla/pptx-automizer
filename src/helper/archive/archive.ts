import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { ArchivedFile, ArchiveType } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';

export default class Archive {
  filename: string;
  buffer: ArchivedFile[] = [];

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

  fromBuffer(relativePath) {
    return this.buffer.find((file) => file.relativePath === relativePath);
  }
}
