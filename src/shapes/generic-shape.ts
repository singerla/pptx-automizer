import { ImportedElement } from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { Shape } from './shape';

export class GenericShape extends Shape {
  sourceElement: XMLDocument;

  constructor(shape: ImportedElement) {
    super(shape);
  }

  async modify(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.replaceIntoSlideTree();
    return this;
  }

  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.appendToSlideTree();
    return this;
  }

  async prepare(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);
    await this.setTargetElement();
    this.applyCallbacks(this.callbacks, this.targetElement);
  }
}
