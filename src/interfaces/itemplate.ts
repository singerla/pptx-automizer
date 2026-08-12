import IArchive from './iarchive';
import { XmlDocument } from '../types/xml-types';
import { AutomizerFile } from '../types/types';

export interface ITemplate {
  location: string;
  file: AutomizerFile;
  archive: IArchive;
  getSlideIdList: () => Promise<XmlDocument>;
}
