import Shape from '../shape'

import { ImportedElement, RootPresTemplate } from '../definitions/app'

export default class Generic extends Shape {
  sourceElement: HTMLElement  

  constructor(shape: ImportedElement) {
    super(shape)
  }
  
  async modify(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.prepare(targetTemplate, targetSlideNumber)
    await this.replaceIntoSlideTree()
    return this
  }

  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.prepare(targetTemplate, targetSlideNumber)
    await this.appendToSlideTree()
    return this
  }

  async prepare(targetTemplate: RootPresTemplate, targetSlideNumber: number) {
    await this.setTarget(targetTemplate, targetSlideNumber)
    this.setTargetElement()
    this.applyCallbacks(this.callbacks, this.targetElement)
  }
}
