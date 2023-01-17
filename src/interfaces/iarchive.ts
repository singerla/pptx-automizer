import ArchiveJszip from '../helper/archive/archive-jszip';
import { AutomizerParams } from '../types/types';
import { InputType } from 'jszip';
import { XmlDocument } from '../types/xml-types';
import ArchiveFs from '../helper/archive/archive-fs';

export type ArchivedFile = {
  name: string;
  relativePath: string;
  content?: XmlDocument;
};

export type ArchiveMode = 'jszip' | 'fs';
export type ArchiveType = ArchiveJszip | ArchiveFs;
export type ArchivedFolderCallback = (file: ArchivedFile) => boolean;
export type ArchiveInput = InputType;

export default interface IArchive {
  read: (file: string, type) => Promise<string | Buffer>;
  write: (file: string, data: string | Buffer) => Promise<ArchiveType>;
  readXml: (file: string) => Promise<XmlDocument>;
  writeXml: (file: string, XmlDocument: XmlDocument) => void;
  folder: (folder: string) => Promise<ArchivedFile[]>;
  fileExists: (file: string) => boolean;
  extract: (file: string) => Promise<ArchiveType>;
  remove: (file: string) => Promise<void>;
  output: (location: string, params: AutomizerParams) => Promise<void>;
  getContent?: (params: AutomizerParams) => Promise<Buffer>;
}
