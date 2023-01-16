import IArchive, { ArchiveInput } from './iarchive';
import { XmlDocument } from '../types/xml-types';

export interface ITemplate {
  location: string;
  file: ArchiveInput;
  archive: IArchive;
  getSlideIdList: () => Promise<XmlDocument>;
}
