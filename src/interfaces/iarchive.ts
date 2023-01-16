import ArchiveJszip from '../helper/archive/archive-jszip';
import { AutomizerParams } from '../types/types';
import JSZip from 'jszip';

export type ArchivedFile = {
  name: string;
  relativePath: string;
};

export type ArchiveType = ArchiveJszip;
export type ArchivedFolderCallback = (file: ArchivedFile) => boolean;

export default interface IArchive {
  read: (file: string, type) => Promise<string | Buffer>;
  folder: (folder: string) => ArchivedFile[];
  count: (pattern: RegExp) => Promise<number>;
  fileExists: (file: string) => boolean;
  extract: (file: string) => Promise<ArchiveType>;
  write: (file: string, data: string | Buffer) => ArchiveType;
  remove: (file: string) => ArchiveType;
  output: (location: string, params: AutomizerParams) => Promise<void>;
}
