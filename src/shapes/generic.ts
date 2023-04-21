import { ImportedElement, ShapeTargetType } from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { Shape } from '../classes/shape';
import { XmlDocument } from '../types/xml-types';

export class GenericShape extends Shape {
  sourceElement: XmlDocument;

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.replaceIntoSlideTree();
    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.appendToSlideTree();
    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);
    await this.setTargetElement();
    this.applyCallbacks(this.callbacks, this.targetElement);
  }
}
