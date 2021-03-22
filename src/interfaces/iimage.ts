import { IShape } from './ishape';
import { RootPresTemplate } from './root-pres-template';

export interface IImage extends IShape {
  sourceFile: string;
  targetFile: string;
  contentTypeMap: any;

  append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<IImage>;
}
