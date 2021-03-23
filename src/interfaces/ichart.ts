import { IShape } from './ishape';
import { RootPresTemplate } from './root-pres-template';

export interface IChart extends IShape {
  sourceNumber: number;
  targetNumber: number;

  append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<IChart>;
}
