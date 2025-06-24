import ArchiveJszip from '../helper/archive/archive-jszip';
import { ArchiveParams, AutomizerFile, AutomizerParams } from '../types/types';
import JSZip, { InputType } from 'jszip';
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
  filename: AutomizerFile;
  params: ArchiveParams;
  read: (
    file: string,
    type: 'string' | 'nodebuffer',
  ) => Promise<string | Buffer>;
  write: (file: string, data: string | Buffer) => Promise<ArchiveType>;
  readXml: (file: string) => Promise<XmlDocument>;
  writeXml: (file: string, XmlDocument: XmlDocument) => void;
  folder: (folder: string) => Promise<ArchivedFile[]>;
  fileExists: (file: string) => boolean;
  extract: (file: string) => Promise<ArchiveType>;
  remove: (file: string) => Promise<void>;
  output: (location: string, params: AutomizerParams) => Promise<void>;
  getContent?: (params: AutomizerParams) => Promise<Buffer>;
  getArchive?: (params: AutomizerParams) => Promise<Buffer>;
  stream?: (
    params: AutomizerParams,
    options: JSZip.JSZipGeneratorOptions<'nodebuffer'>,
  ) => Promise<NodeJS.ReadableStream>;
  getFinalArchive?: () => Promise<JSZip>;
}
