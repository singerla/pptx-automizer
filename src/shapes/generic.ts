import Shape from '../shape'

import { ImportedElement, RootPresTemplate } from '../definitions/app'

export default class Generic extends Shape {
  sourceElement: HTMLElement  

  constructor(shape: ImportedElement) {
    super(shape)
    this.sourceElement = shape.element
  }
  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.setTarget(targetTemplate, targetSlideNumber)
    this.targetElement = <HTMLElement> this.sourceElement.cloneNode(true)
    this.applyCallbacks(this.callbacks, this.targetElement)
    await this.appendToSlideTree()
    return this
  }
}
