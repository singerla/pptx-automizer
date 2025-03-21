import { AutomizerParams, StatusTracker } from '../types/types';
import { PresTemplate } from './pres-template';
import { RootPresTemplate } from './root-pres-template';
import { ContentTracker } from '../helper/content-tracker';
import { ModifyXmlCallback } from '../types/xml-types';
import Automizer from '../automizer';

export interface IPresentationProps {
  rootTemplate: RootPresTemplate;
  templates: PresTemplate[];
  params: AutomizerParams;
  timer: number;
  status?: StatusTracker;
  content?: ContentTracker;

  getTemplate(name: string): PresTemplate;

  modify(cb: ModifyXmlCallback): Automizer;
}
