import JSZip from 'jszip';

export interface ITemplate {
  location: string;
  file: Promise<Buffer>;
  archive: Promise<JSZip>;
}
