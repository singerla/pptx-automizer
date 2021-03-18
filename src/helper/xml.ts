import { DOMParser, XMLSerializer } from 'xmldom'
import FileHelper from './file'
import JSZip from 'jszip'

import { Target } from '../definitions/app'
import { DefaultAttribute, OverrideAttribute, RelationshipAttribute, XMLElement } from '../definitions/xml'
import { TargetByRelIdMap } from '../definitions/constants'

export default class XmlHelper {

  static async getXmlFromArchive(archive: JSZip, file: string): Promise<Document> {
    let xmlDocument = await FileHelper.extractFromArchive(archive, file)
    const dom = new DOMParser()
    return dom.parseFromString(xmlDocument)
  }
  
  static async writeXmlToArchive(archive: JSZip, file: string, xml: any): Promise<JSZip> {
    let s = new XMLSerializer()
    let xmlBuffer = s.serializeToString(xml)
    
    return archive.file(file, xmlBuffer)
  }

  static async appendIf(element: XMLElement): Promise<HTMLElement | boolean> {
    let xml = await XmlHelper.getXmlFromArchive(element.archive, element.file)
    
    if(element.clause !== undefined) {
      if(!element.clause(xml)) return false
    }

    return XmlHelper.append(element)
  }

  static async append(element: XMLElement): Promise<HTMLElement> {
    let xml = await XmlHelper.getXmlFromArchive(element.archive, element.file)
    
    let newElement = xml.createElement(element.tag)
    for(let attribute in element.attributes) {
      let value = element.attributes[attribute]
      newElement.setAttribute(attribute, (typeof value === 'function') ? value(xml) : value)
    }

    let parent = element.parent(xml)
    parent.appendChild(newElement)

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml)

    return newElement
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

  static async getTargetsFromRelationships(archive: JSZip, path: string, prefix: string, suffix?: string | RegExp): Promise<Target[]>{
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      let target = element.getAttribute('Target')
      if(target.indexOf(prefix) === 0) {
        rels.push(<Target>{
          file: target,
          rId: element.getAttribute('Id'),
          number: Number(target.replace(prefix, '').replace(suffix || '.xml', ''))
        })
      }
    })
  }

  static async getTargetsByRelationshipType(archive: JSZip, path: string, type: string): Promise<Target[]>{
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      let target = element.getAttribute('Type')
      if(target === type) {
        rels.push(<Target>{
          file: element.getAttribute('Target'),
          rId: element.getAttribute('Id'),
        })
      }
    })
  }

  static async getRelationships(archive: JSZip, path: string, cb: Function) {
    let xml = await XmlHelper.getXmlFromArchive(archive, path)
    let relationships = xml.getElementsByTagName('Relationship')
    let rels = []
    for(let i in relationships) {
      let element = relationships[i]
      if(element.getAttribute !== undefined) {
        cb(element, rels)
      }
    }
    return rels
  }

  static findByAttribute(xml: HTMLElement | Document, tagName: string, attributeName: string, attributeValue: string): Boolean {
    let elements = xml.getElementsByTagName(tagName)
    for(let i in elements) {
      let element = elements[i]
      if(element.getAttribute !== undefined) {
        if(element.getAttribute(attributeName) === attributeValue) {
          return true
        }
      }
    }
    return false
  }

  static async replaceAttribute(archive: JSZip, path: string, tagName: string, attributeName: string, attributeValue: string, replaceValue: string): Promise<JSZip> {
    let xml = await XmlHelper.getXmlFromArchive(archive, path)
    let elements = xml.getElementsByTagName(tagName)
    for(let i in elements) {
      let element = elements[i]
      if(element.getAttribute !== undefined && element.getAttribute(attributeName) === attributeValue) {
        element.setAttribute(attributeName, replaceValue)
      }
    }
    return XmlHelper.writeXmlToArchive(archive, path, xml)
  }

  static async getTargetByRelId(archive: JSZip, slideNumber: number, element: HTMLElement, type: string): Promise<Target> {
    let params = TargetByRelIdMap[type]
    let sourceRid = element.getElementsByTagName(params.relRootTag)[0].getAttribute(params.relAttribute)
    let relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`
    let imageRels = await XmlHelper.getTargetsFromRelationships(archive, relsPath, params.prefix, params.expression)
    let target = imageRels.find(rel => rel.rId === sourceRid)

    return target
  }

  static async findByElementName(archive: JSZip, path: string, name: string): Promise<any> {
    let slideXml = await XmlHelper.getXmlFromArchive(archive, path)
    let names = slideXml.getElementsByTagName('p:cNvPr')
    
    for(let i in names) {
      if(names[i].getAttribute && names[i].getAttribute('name') === name) {
        return names[i].parentNode.parentNode
      }
    }

    return null
  }

  static findByAttributeValue(nodes: any, attributeName: string, attributeValue:string): HTMLElement {
    for(let i in nodes) {
      if(nodes[i].getAttribute && nodes[i].getAttribute(attributeName) === attributeValue) {
        return nodes[i]
      }
    }
    return null
  }

  static createContentTypeChild(archive: JSZip, attributes: OverrideAttribute | DefaultAttribute): XMLElement {
    return {
      archive: archive,
      file: `[Content_Types].xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes: attributes
    }
  }

  static createRelationshipChild(archive: JSZip, targetRelFile:string, attributes: RelationshipAttribute): XMLElement {
    return {
      archive: archive,
      file: targetRelFile,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: attributes
    }
  }

  static setChartData(chart, workbook, data) {
    let series = chart.getElementsByTagName('c:ser')

    for(let c in data.categories) {
      for(let s in data.categories[c].values) {
        series[s].getElementsByTagName('c:cat')[0]
          .getElementsByTagName('c:v')[c]
          .firstChild.data = data.categories[c].label

        series[s].getElementsByTagName('c:v')[0]
          .firstChild.data = data.series[s].label

        series[s].getElementsByTagName('c:val')[0]
          .getElementsByTagName('c:v')[c]
          .firstChild.data = String(data.categories[c].values[s])
      }
    }

    XmlHelper.setWorkbookData(workbook, data)
  }

  static setWorkbookData(workbook, data) {
    let rows = workbook.sheet.getElementsByTagName('row')
  
    for(let c in data.categories) {
      let r = Number(c) + 1
      let stringId = XmlHelper.appendSharedString(workbook.sharedStrings, data.categories[c].label)
      let rowLabel = rows[r].getElementsByTagName('c')[0].getElementsByTagName('v')[0]
      rowLabel.firstChild.data = String(stringId)
  
      for(let s in data.categories[c].values) {
        let v = Number(s) + 1
        rows[r].getElementsByTagName('c')[v]
          .getElementsByTagName('v')[0]
          .firstChild.data = String(data.categories[c].values[s])
      }
    }
  
    for(let s in data.series) {
      let c = Number(s) + 1
      let colLabel = rows[0].getElementsByTagName('c')[c].getElementsByTagName('v')[0]
      let stringId = XmlHelper.appendSharedString(workbook.sharedStrings, data.series[s].label)
      
      colLabel.firstChild.data = String(stringId)
  
      workbook.table.getElementsByTagName('tableColumn')[c].setAttribute('name', data.series[s].label)
    }
  }

  static appendSharedString(sharedStrings: Document, string: string): number {
    let strings = sharedStrings.getElementsByTagName('sst')[0]
    let newLabel = sharedStrings.createTextNode(string)
    let newText = sharedStrings.createElement('t')
    newText.appendChild(newLabel)

    let newString = sharedStrings.createElement('si')
    newString.appendChild(newText)

    strings.appendChild(newString)
    
    let stringId = strings.getElementsByTagName('si').length - 1
    return stringId
  }
}
