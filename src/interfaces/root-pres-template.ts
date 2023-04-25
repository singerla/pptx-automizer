import { ISlide } from './islide';
import { ICounter } from './icounter';
import { ITemplate } from './itemplate';
import { ContentTracker } from '../helper/content-tracker';
import { IMaster } from './imaster';
import Automizer from '../automizer';

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  masters: IMaster[];
  counter: ICounter[];
  mapContents: (
    type: string,
    key: string,
    sourceId: number,
    targetId: number,
  ) => void;
  getMappedContent: (type: string, key: string, sourceId: number) => any;
  count(name: string): number;
  incrementCounter(name: string): number;
  appendSlide(slide: ISlide): Promise<void>;
  appendMasterSlide(slideMaster: IMaster): Promise<void>;
  countExistingSlides(): Promise<void>;
  truncate(): Promise<void>;
  content?: ContentTracker;
  automizer?: Automizer;
}
