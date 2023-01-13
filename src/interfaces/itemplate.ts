import { InputType } from 'jszip';
import { FileProxy } from '../helper/file-proxy';

export interface ITemplate {
  location: string;
  file: InputType;
  archive: FileProxy;
  getSlideIdList: () => Promise<Document>;
}
