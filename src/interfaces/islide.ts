import JSZip from 'jszip';
import { RootPresTemplate } from './root-pres-template';
import { SlideModificationCallback } from '../types/types';

export interface ISlide {
  sourceArchive: JSZip;
  sourceNumber: number;
  modifications: SlideModificationCallback[];

  modify(callback: SlideModificationCallback): void;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: string): void;
}
