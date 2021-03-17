import JSZip from 'jszip'
import Shape from '../shape'

import { ImportedElement, RootPresTemplate } from '../definitions/app'
import GeneralHelper from '../helper/general'

export default class Generic extends Shape {
  sourceElement: HTMLElement  
  callbacks: any

  constructor(info: ImportedElement, sourceArchive: JSZip, sourceSlideNumber?:number) {
    let relsXmlInfo = {
      file: null, number: null
    }
    super(relsXmlInfo, sourceArchive, sourceSlideNumber)
    this.sourceElement = info.element
    this.callbacks = GeneralHelper.arrayify(info.callback)
  }
  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Generic> {
    await this.setTarget(targetTemplate, targetSlideNumber)
    this.targetElement = <HTMLElement> this.sourceElement.cloneNode(true)
    this.applyCallbacks(this.callbacks, this.targetElement)
    await this.appendToSlideTree()
    
    return this
  }
}
