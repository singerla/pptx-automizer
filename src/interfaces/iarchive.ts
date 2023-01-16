import ArchiveJszip from '../helper/archive/archive-jszip';
import { AutomizerParams } from '../types/types';
import { InputType } from 'jszip';
import { XmlDocument } from '../types/xml-types';

export type ArchivedFile = {
  name: string;
  relativePath: string;
  content?: XmlDocument;
};

export type ArchiveType = ArchiveJszip;
export type ArchivedFolderCallback = (file: ArchivedFile) => boolean;
export type ArchiveInput = InputType;

export default interface IArchive {
  read: (file: string, type) => Promise<string | Buffer>;
  write: (file: string, data: string | Buffer) => ArchiveType;
  readXml: (file: string) => Promise<XmlDocument>;
  writeXml: (file: string, XmlDocument: XmlDocument) => void;
  folder: (folder: string) => ArchivedFile[];
  count: (pattern: RegExp) => Promise<number>;
  fileExists: (file: string) => boolean;
  extract: (file: string) => Promise<ArchiveType>;

  remove: (file: string) => ArchiveType;
  getContent: (params: AutomizerParams) => Promise<Buffer>;
  output: (location: string, params: AutomizerParams) => Promise<void>;
}
