import Shape from '../classes/shape';

import { ImportedElement } from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';

export default class Generic extends Shape {
  sourceElement: HTMLElement;

  constructor(shape: ImportedElement) {
    super(shape);
  }

  async modify(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.replaceIntoSlideTree();
    return this;
  }

  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.appendToSlideTree();
    return this;
  }

  async prepare(targetTemplate: RootPresTemplate, targetSlideNumber: number) {
    await this.setTarget(targetTemplate, targetSlideNumber);
    this.setTargetElement();
    this.applyCallbacks(this.callbacks, this.targetElement);
  }
}
