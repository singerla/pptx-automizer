import { ISlide } from './islide';
import { ICounter } from './icounter';
import { ITemplate } from './itemplate';
import { ContentTracker } from '../helper/content-tracker';

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  counter: ICounter[];

  count(name: string): number;

  incrementCounter(name: string): number;

  appendSlide(slide: ISlide): Promise<void>;
  countExistingSlides(): Promise<void>;
  truncate(): Promise<void>;
  content?: ContentTracker;
}
