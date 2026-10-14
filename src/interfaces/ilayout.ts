import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  ShapeModificationCallback,
  SlideModificationCallback,
} from '../types/types';
import IArchive from './iarchive';
import { ModifyXmlCallback } from '../types/xml-types';

export interface ILayout {
  sourceArchive: IArchive;
  sourceNumber: number;
  modify(callback: SlideModificationCallback): void;
  modifyRelations(callback: ModifyXmlCallback): void;
  append(targetTemplate: RootPresTemplate): Promise<void>;
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): ILayout;
  modifyElement(
    selector: FindElementSelector,
    callback: ShapeModificationCallback | ShapeModificationCallback[],
  ): ILayout;
  removeElement(selector: FindElementSelector): ILayout;
}
