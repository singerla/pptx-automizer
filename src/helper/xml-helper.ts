import JSZip from 'jszip';

import { DOMParser, XMLSerializer } from 'xmldom';
import { FileHelper } from './file-helper';
import { DefaultAttribute, OverrideAttribute, RelationshipAttribute, XMLElement } from '../types/xml-types';
import { TargetByRelIdMap } from '../constants/constants';
import { XmlPrettyPrint } from './xml-pretty-print';
import { Target } from '../types/types';

export class XmlHelper {

  static async getXmlFromArchive(archive: JSZip, file: string): Promise<Document> {
    const xmlDocument = await FileHelper.extractFromArchive(archive, file);
    const dom = new DOMParser();
    return dom.parseFromString(xmlDocument);
  }

  static async writeXmlToArchive(archive: JSZip, file: string, xml: any): Promise<JSZip> {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(xml);

    return archive.file(file, xmlBuffer);
  }

  static async appendIf(element: XMLElement): Promise<HTMLElement | boolean> {
    const xml = await XmlHelper.getXmlFromArchive(element.archive, element.file);

    return element.clause !== undefined && !element.clause(xml)
      ? false
      : XmlHelper.append(element);

  }

  static async append(element: XMLElement): Promise<HTMLElement> {
    const xml = await XmlHelper.getXmlFromArchive(element.archive, element.file);

    const newElement = xml.createElement(element.tag);
    for (const attribute in element.attributes) {
      const value = element.attributes[attribute];
      newElement.setAttribute(attribute, (typeof value === 'function') ? value(xml) : value);
    }

    const parent = element.parent(xml);
    parent.appendChild(newElement);

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return newElement;
  }

  static async getNextRelId(rootArchive, file): Promise<string> {
    const presentationRelsXml = await XmlHelper.getXmlFromArchive(rootArchive, file);
    const increment = (max: number) => 'rId' + max;
    const rid = XmlHelper.getMaxId(presentationRelsXml.documentElement.childNodes, 'Id', true);

    return increment(rid);
  }

  static getMaxId(rels: { [x: string]: any }, attribute: string, increment?: boolean): number {
    let max = 0;
    for (const i in rels) {
      const rel = rels[i];
      if (rel.getAttribute !== undefined) {
        const id = Number(rel.getAttribute(attribute).replace('rId', ''));
        max = (id > max) ? id : max;
      }
    }

    switch (typeof increment) {
      case 'boolean' :
        return ++max;
      default:
        return max;
    }
  }

  static async getTargetsFromRelationships(archive: JSZip, path: string, prefix: string, suffix?: string | RegExp): Promise<Target[]> {
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      const target = element.getAttribute('Target');
      if (target.indexOf(prefix) === 0) {
        rels.push({
          file: target,
          rId: element.getAttribute('Id'),
          number: Number(target.replace(prefix, '').replace(suffix || '.xml', ''))
        } as Target);
      }
    });
  }

  static async getTargetsByRelationshipType(archive: JSZip, path: string, type: string): Promise<Target[]> {
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      const target = element.getAttribute('Type');
      if (target === type) {
        rels.push({
          file: element.getAttribute('Target'),
          rId: element.getAttribute('Id'),
        } as Target);
      }
    });
  }

  static async getRelationships(archive: JSZip, path: string, cb: Function) {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const relationships = xml.getElementsByTagName('Relationship');
    const rels = [];

    Object.keys(relationships)
      .map(key => relationships[key] as Element)
      .filter(element => element.getAttribute !== undefined)
      .forEach(element => cb(element, rels));

    return rels;
  }

  static findByAttribute(xml: HTMLElement | Document, tagName: string, attributeName: string, attributeValue: string): boolean {
    const elements = xml.getElementsByTagName(tagName);
    for (const i in elements) {
      const element = elements[i];
      if (element.getAttribute !== undefined) {
        if (element.getAttribute(attributeName) === attributeValue) {
          return true;
        }
      }
    }
    return false;
  }

  static async replaceAttribute(archive: JSZip, path: string, tagName: string, attributeName: string, attributeValue: string, replaceValue: string): Promise<JSZip> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const elements = xml.getElementsByTagName(tagName);
    for (const i in elements) {
      const element = elements[i];
      if (element.getAttribute !== undefined && element.getAttribute(attributeName) === attributeValue) {
        element.setAttribute(attributeName, replaceValue);
      }
    }
    return XmlHelper.writeXmlToArchive(archive, path, xml);
  }

  static async getTargetByRelId(archive: JSZip, slideNumber: number, element: HTMLElement, type: string): Promise<Target> {
    const params = TargetByRelIdMap[type];
    const sourceRid = element.getElementsByTagName(params.relRootTag)[0].getAttribute(params.relAttribute);
    const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
    const imageRels = await XmlHelper.getTargetsFromRelationships(archive, relsPath, params.prefix, params.expression);
    const target = imageRels.find(rel => rel.rId === sourceRid);

    return target;
  }

  static async findByElementName(archive: JSZip, path: string, name: string): Promise<any> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByName(slideXml, name);
  }

  static findByName(doc: Document, name: string): any {
    const names = doc.getElementsByTagName('p:cNvPr');

    for (const i in names) {
      if (names[i].getAttribute && names[i].getAttribute('name') === name) {
        return names[i].parentNode.parentNode;
      }
    }

    return null;
  }

  static findByAttributeValue(nodes: any, attributeName: string, attributeValue: string): HTMLElement {
    for (const i in nodes) {
      if (nodes[i].getAttribute && nodes[i].getAttribute(attributeName) === attributeValue) {
        return nodes[i];
      }
    }
    return null;
  }

  static createContentTypeChild(archive: JSZip, attributes: OverrideAttribute | DefaultAttribute): XMLElement {
    return {
      archive,
      file: `[Content_Types].xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes
    };
  }

  static createRelationshipChild(archive: JSZip, targetRelFile: string, attributes: RelationshipAttribute): XMLElement {
    return {
      archive,
      file: targetRelFile,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes
    };
  }

  static setChartData(chart, workbook, data) {
    const series = chart.getElementsByTagName('c:ser');

    for (const c in data.categories) {
      for (const s in data.categories[c].values) {
        series[s].getElementsByTagName('c:cat')[0]
          .getElementsByTagName('c:v')[c]
          .firstChild.data = data.categories[c].label;

        series[s].getElementsByTagName('c:v')[0]
          .firstChild.data = data.series[s].label;

        series[s].getElementsByTagName('c:val')[0]
          .getElementsByTagName('c:v')[c]
          .firstChild.data = String(data.categories[c].values[s]);
      }
    }

    XmlHelper.setWorkbookData(workbook, data);
  }

  static setWorkbookData(workbook, data) {
    const rows = workbook.sheet.getElementsByTagName('row');

    for (const c in data.categories) {
      const r = Number(c) + 1;
      const stringId = XmlHelper.appendSharedString(workbook.sharedStrings, data.categories[c].label);
      const rowLabel = rows[r].getElementsByTagName('c')[0].getElementsByTagName('v')[0];
      rowLabel.firstChild.data = String(stringId);

      for (const s in data.categories[c].values) {
        const v = Number(s) + 1;
        rows[r].getElementsByTagName('c')[v]
          .getElementsByTagName('v')[0]
          .firstChild.data = String(data.categories[c].values[s]);
      }
    }

    for (const s in data.series) {
      const c = Number(s) + 1;
      const colLabel = rows[0].getElementsByTagName('c')[c].getElementsByTagName('v')[0];
      const stringId = XmlHelper.appendSharedString(workbook.sharedStrings, data.series[s].label);

      colLabel.firstChild.data = String(stringId);

      workbook.table.getElementsByTagName('tableColumn')[c].setAttribute('name', data.series[s].label);
    }
  }

  static appendSharedString(sharedStrings: Document, stringValue: string): number {
    const strings = sharedStrings.getElementsByTagName('sst')[0];
    const newLabel = sharedStrings.createTextNode(stringValue);
    const newText = sharedStrings.createElement('t');
    newText.appendChild(newLabel);

    const newString = sharedStrings.createElement('si');
    newString.appendChild(newText);

    strings.appendChild(newString);

    return strings.getElementsByTagName('si').length - 1;
  }

  static dump(element: any) {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(element);

    const p = new XmlPrettyPrint(xmlBuffer);
    p.dump();
  }

}
