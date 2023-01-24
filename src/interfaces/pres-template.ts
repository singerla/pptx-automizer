import { ITemplate } from './itemplate';
import { SlideInfo } from '../types/xml-types';

export interface PresTemplate extends ITemplate {
  name: string;
  setCreationIds(): Promise<SlideInfo[]>;
  creationIds?: SlideInfo[];
  useCreationIds?: boolean;
}
