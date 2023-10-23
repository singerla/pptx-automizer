import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  ModificationCallback,
  SlideModificationCallback,
  SourceIdentifier,
} from '../types/types';
import IArchive from './iarchive';

export interface ISlide {
  sourceArchive: IArchive;
  sourceNumber: SourceIdentifier;
  modify(callback: SlideModificationCallback): void;
  append(targetTemplate: RootPresTemplate): Promise<void>;
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ModificationCallback | ModificationCallback[],
  ): ISlide;
  modifyElement(
    selector: FindElementSelector,
    callback: ModificationCallback | ModificationCallback[],
  ): ISlide;
  removeElement(selector: FindElementSelector): ISlide;
  useSlideLayout(targetLayout?: number | string): ISlide;
  getAllTextElementIds(): Promise<string[]>;
}
