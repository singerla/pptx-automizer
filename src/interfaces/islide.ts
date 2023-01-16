import { RootPresTemplate } from './root-pres-template';
import {
  SlideModificationCallback,
  SourceSlideIdentifier,
} from '../types/types';
import IArchive from './iarchive';

export interface ISlide {
  sourceArchive: IArchive;
  sourceNumber: SourceSlideIdentifier;
  modifications: SlideModificationCallback[];

  modify(callback: SlideModificationCallback): void;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: string): void;
}
