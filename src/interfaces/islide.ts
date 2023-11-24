import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  GenerateElements,
  GenerateOnSlideCallback,
  ModificationCallback,
  SlideModificationCallback,
  SourceIdentifier,
} from '../types/types';
import IArchive from './iarchive';
import { ElementInfo } from '../types/xml-types';

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

  /**
   * Use PptxGenJs to generate a new element on this slide.
   * @param generate
   * @param objectName
   */
  generate(generate: GenerateOnSlideCallback, objectName?: string): ISlide;
  removeElement(selector: FindElementSelector): ISlide;
  useSlideLayout(targetLayout?: number | string): ISlide;
  getAllElements(filterTags?: string[]): Promise<ElementInfo[]>;
  getAllTextElementIds(): Promise<string[]>;
  getGeneratedElements(): GenerateElements[];
}
