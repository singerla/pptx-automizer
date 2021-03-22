import JSZip, { InputType } from 'jszip';

export interface ITemplate {
  location: string;
  file: InputType;
  archive: Promise<JSZip>;
}
