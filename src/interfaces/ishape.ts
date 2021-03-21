import JSZip from 'jszip';

export interface IShape {
  sourceArchive: JSZip;
  targetArchive: JSZip;
}
