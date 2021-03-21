import JSZip from 'jszip';
import { RootPresTemplate } from './root-pres-template';

export interface ISlide {
  sourceArchive: JSZip;
  sourceNumber: number;
  modifications: Function[];
  modify: Function;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: Function | string): void;
}
