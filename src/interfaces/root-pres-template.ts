import { ISlide } from './islide';
import { ICounter } from './icounter';
import { ITemplate } from './itemplate';
import { ContentTracker } from '../helper/content-tracker';
import Automizer from '../automizer';
import { IMaster } from './imaster';
import { ILayout } from './ilayout';
import { MediaFile } from '../types/types';

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  masters: IMaster[];
  counter: ICounter[];
  mediaFiles: MediaFile[];
  mapContents: (
    type: string,
    key: string,
    sourceId: number,
    targetId: number,
    name?: string,
  ) => void;
  getMappedContent: (type: string, key: string, sourceId: number) => any;
  getNamedMappedContent: (type: string, name: string) => any;

  count(name: string): number;

  incrementCounter(name: string): number;

  appendSlide(slide: ISlide): Promise<void>;

  appendMasterSlide(slideMaster: IMaster): Promise<void>;

  appendLayout(slideLayout: ILayout): Promise<void>;

  runExternalGenerator?(): Promise<void>;

  cleanupExternalGenerator?(): Promise<void>;

  countExistingSlides(): Promise<void>;

  truncate(): Promise<void>;

  content?: ContentTracker;
  automizer?: Automizer;
}
