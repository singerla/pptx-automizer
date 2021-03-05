import { DOMParser, XMLSerializer } from 'xmldom'
import { IArchive } from '../types/interfaces'
import FileHelper from './file'
import { XMLElement } from '../types/xml'

export default class XmlHelper {

  static getXmlFromArchive(archive: IArchive, file: string) {
    return FileHelper.extractFromArchive(archive, file)
      .then(XmlHelper.parseXmlDocument)
  }
  
  static parseXmlDocument(xmlDocument: any) {
    const dom = new DOMParser()
    return dom.parseFromString(xmlDocument)
  }

  static async writeXmlToArchive(archive: IArchive, file: string, xml: any) {
    let s = new XMLSerializer()
    let xmlBuffer = s.serializeToString(xml)
    
    return archive.file(file, xmlBuffer)
  }

  static async append(element: XMLElement): Promise<void> {
    let xml = await XmlHelper.getXmlFromArchive(element.archive, element.file)

    let newElement = xml.createElement(element.tag)
    for(let attribute in element.attributes) {
      let value = element.attributes[attribute]
      newElement.setAttribute(attribute, (typeof value === 'function') ? value(xml) : value)
    }

    let parent = element.parent(xml)
    parent.appendChild(newElement)

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml)
  }

  static async getNextRelId(rootArchive, file): Promise<string> {
    let presentationRelsXml = await XmlHelper.getXmlFromArchive(rootArchive, file)
    let increment: Function = (max: number) => 'rId' + max
    let rid = XmlHelper.getMaxId(presentationRelsXml.documentElement.childNodes, 'Id', true)

    return increment(rid)
  }

  static getMaxId(rels: { [x: string]: any }, attribute: string, increment?: boolean): number {
    let max = 0
    for(let i in rels) {
      let rel = rels[i]
      if(rel.getAttribute !== undefined) {
        let id = Number(rel.getAttribute(attribute).replace('rId', ''))
        max = (id > max) ? id : max
      }
    }

    switch(typeof increment) {
      case 'boolean' : return ++max
      default: return max
    }
  }

}