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

  modifyRelations(callback: SlideModificationCallback): void;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  /**
   * Add an element from one of the loaded templates to this slide.
   * @param presName Filename or custom name of source template
   * @param slideNumber The slide number or slide creationID where the source element can be found on
   * @param selector The name or element creationID of the source element
   * @param callback You can pass a callback to modify the element on the target slide after insertion.
   */
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ModificationCallback | ModificationCallback[],
  ): ISlide;

  /**
   * Modify an element on the current slide.
   * @param selector The name or element creationID of the element on current slide
   * @param callback Pass a callback to modify the element.
   */
  modifyElement(
    selector: FindElementSelector,
    callback: ModificationCallback | ModificationCallback[],
  ): ISlide;

  /**
   * Use PptxGenJs to generate a new element from scratch on this slide.
   * @param generate Pass a callback to create an element on current slide
   * @param objectName Give a custom name for this element or automizer will create a random uuid
   */
  generate(generate: GenerateOnSlideCallback, objectName?: string): ISlide;

  /**
   * You can remove an element on the current slide.
   * @param selector Pass element name or creationId to find target element.
   */
  removeElement(selector: FindElementSelector): ISlide;

  useSlideLayout(targetLayout?: number | string): ISlide;

  getElement(selector: FindElementSelector): Promise<ElementInfo>;

  getAllElements(filterTags?: string[]): Promise<ElementInfo[]>;

  getAllTextElementIds(): Promise<string[]>;

  getGeneratedElements(): GenerateElements[];

  /**
   * Asynchronously retrieves the dimensions of a slide.
   */
  getDimensions(): Promise<{ width: number; height: number }>;

  /**
   * Remove a slide from output. The slide will be calculated, but
   * eventually withdrawn from slide list.
   * Slide number starts by 1.
   */
  remove(number: number): void;
}
