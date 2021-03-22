import { ISlide } from './islide';
import { ICounter } from './icounter';
import { ITemplate } from './itemplate';

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  counter: ICounter[];

  count(name: string): number;

  incrementCounter(name: string): number;

  appendSlide(slide: ISlide): Promise<void>;
}
