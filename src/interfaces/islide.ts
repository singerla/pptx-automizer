import JSZip from 'jszip';
import { RootPresTemplate } from './root-pres-template';
import {
  SlideModificationCallback,
  SourceSlideIdentifier,
} from '../types/types';
import { FileProxy } from '../helper/file-proxy';

export interface ISlide {
  sourceArchive: FileProxy;
  sourceNumber: SourceSlideIdentifier;
  modifications: SlideModificationCallback[];

  modify(callback: SlideModificationCallback): void;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: string): void;
}
