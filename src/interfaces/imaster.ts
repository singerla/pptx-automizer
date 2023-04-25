import { RootPresTemplate } from './root-pres-template';
import { SlideModificationCallback, SourceIdentifier } from '../types/types';
import IArchive from './iarchive';

export interface IMaster {
  sourceArchive: IArchive;
  sourceNumber: SourceIdentifier;
  modifications: SlideModificationCallback[];

  append(targetTemplate: RootPresTemplate): Promise<void>;

  // modify(callback: SlideModificationCallback): void;
  // addElement(presName: string, slideNumber: number, selector: string): void;
}
