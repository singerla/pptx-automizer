import { RootPresTemplate } from './root-pres-template';
import {
  FindElementSelector,
  ShapeModificationCallback,
  SlideModificationCallback,
  SourceIdentifier,
} from '../types/types';
import IArchive from './iarchive';

export interface ILayout {
  sourceArchive: IArchive;
  sourceNumber: number;
  modify(callback: SlideModificationCallback): void;
  append(targetTemplate: RootPresTemplate): Promise<void>;
}
