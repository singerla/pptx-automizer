import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  ShapeModificationCallback,
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
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): ISlide;
  modifyElement(
    selector: FindElementSelector,
    callback: ShapeModificationCallback | ShapeModificationCallback[],
  ): ISlide;
  removeElement(selector: FindElementSelector): ISlide;
  useSlideLayout(index?: number): ISlide;
}
