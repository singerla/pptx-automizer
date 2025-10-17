import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  ShapeModificationCallback,
  SlideModificationCallback,
  SourceIdentifier,
} from '../types/types';
import IArchive from './iarchive';
import { ModifyXmlCallback } from '../types/xml-types';

export interface IMaster {
  sourceArchive: IArchive;
  sourceNumber: number;
  key: string;
  modify(callback: SlideModificationCallback): void;
  modifyRelations(callback: ModifyXmlCallback): void;
  append(targetTemplate: RootPresTemplate): Promise<void>;
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): IMaster;
  modifyElement(
    selector: FindElementSelector,
    callback: ShapeModificationCallback | ShapeModificationCallback[],
  ): IMaster;
  removeElement(selector: FindElementSelector): IMaster;
}
