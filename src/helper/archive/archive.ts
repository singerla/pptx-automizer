import { DOMParser, Node, XMLSerializer } from '@xmldom/xmldom';
import { ArchivedFile, ArchiveType } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import {
  ArchiveParams,
  AutomizerFile,
  AutomizerParams,
} from '../../types/types';
import JSZip from 'jszip';
import { patchXmldomSaxRegExpCache } from '../xmldom-sax-patch';

patchXmldomSaxRegExpCache();

export default class Archive {
  filename: AutomizerFile;
  params: ArchiveParams;
  buffer: Map<string, ArchivedFile> = new Map();
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

  serializeXml(xml: XmlDocument | Node) {
    const s = new XMLSerializer();
    return s.serializeToString(<Node>xml);
  }

  async writeBuffer(archiveType: ArchiveType) {
    for (const buffered of this.buffer.values()) {
      const serialized = this.serializeXml(buffered.content);
      await archiveType.write(buffered.relativePath, serialized);
    }
  }

  toBuffer(relativePath: string, content: XmlDocument): void {
    this.buffer.set(relativePath, {
      relativePath,
      name: relativePath,
      content: content,
    });
  }

  /**
   * Serializes a buffered part back into the underlying archive and drops
   * its DOM from the buffer. A parsed xmldom document costs ~25x its XML
   * source size, so parts that are finished (an appended slide after
   * cleanSlide, a master/layout after append) must not stay buffered for
   * the rest of the run. Re-reading a flushed part re-parses it from the
   * serialized content written here. No-op if the part is not buffered.
   */
  protected async flushBuffered(
    archiveType: ArchiveType,
    relativePath: string,
  ): Promise<void> {
    const buffered = this.buffer.get(relativePath);
    if (!buffered) {
      return;
    }
    await archiveType.write(relativePath, this.serializeXml(buffered.content));
    this.buffer.delete(relativePath);
  }

  setOptions(params: AutomizerParams): void {
    if (params.compression > 0) {
      this.options.compression = 'DEFLATE';
      this.options.compressionOptions = {
        level: params.compression,
      };
    }
  }

  fromBuffer(relativePath: string): ArchivedFile | undefined {
    return this.buffer.get(relativePath);
  }
}
