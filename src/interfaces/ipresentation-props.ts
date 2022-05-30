import { AutomizerParams, StatusTracker } from '../types/types';
import { PresTemplate } from './pres-template';
import { RootPresTemplate } from './root-pres-template';

export interface IPresentationProps {
  rootTemplate: RootPresTemplate;
  templates: PresTemplate[];
  params: AutomizerParams;
  timer: number;
  status?: StatusTracker;
  getTemplate(name: string): PresTemplate;
}
