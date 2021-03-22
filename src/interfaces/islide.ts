import JSZip from 'jszip';
import { RootPresTemplate } from './root-pres-template';
import { ModificationCallback } from '../types/types';

export interface ISlide {
  sourceArchive: JSZip;
  sourceNumber: number;
  modifications: ModificationCallback[];

  modify(callback: ModificationCallback): void;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: string): void;
}
