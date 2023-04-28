import { ITemplate } from './itemplate';
import { SlideInfo } from '../types/xml-types';

export interface PresTemplate extends ITemplate {
  name: string;
  setCreationIds(): Promise<SlideInfo[]>;
  getAllSlideNumbers(): Promise<number[]>;
  creationIds?: SlideInfo[];
  useCreationIds?: boolean;
}
