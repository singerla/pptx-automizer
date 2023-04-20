import { ISlide } from './islide';
import { ICounter } from './icounter';
import { ITemplate } from './itemplate';
import { ContentTracker } from '../helper/content-tracker';
import { IMaster } from './imaster';

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  masters: IMaster[];
  counter: ICounter[];
  count(name: string): number;
  incrementCounter(name: string): number;
  appendSlide(slide: ISlide): Promise<void>;
  appendMasterSlide(slideMaster: IMaster): Promise<void>;
  countExistingSlides(): Promise<void>;
  truncate(): Promise<void>;
  content?: ContentTracker;
}
