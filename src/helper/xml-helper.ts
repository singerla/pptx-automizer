import JSZip from 'jszip';

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { FileHelper } from './file-helper';
import {
  DefaultAttribute,
  OverrideAttribute,
  RelationshipAttribute,
  HelperElement,
} from '../types/xml-types';
import { TargetByRelIdMap } from '../constants/constants';
import { XmlPrettyPrint } from './xml-pretty-print';
import {
  GetRelationshipsCallback,
  SourceSlideIdentifier,
  Target,
} from '../types/types';
import { XmlTemplateHelper } from './xml-template-helper';
import { vd } from './general-helper';

export class XmlHelper {
  static async getXmlFromArchive(
    archive: JSZip,
    file: string,
  ): Promise<XMLDocument> {
    const xmlDocument = (await FileHelper.extractFromArchive(
      archive,
      file,
    )) as string;
    const dom = new DOMParser();
    return dom.parseFromString(xmlDocument);
  }

  static async writeXmlToArchive(
    archive: JSZip,
    file: string,
    xml: XMLDocument,
  ): Promise<JSZip> {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(xml);

    return archive.file(file, xmlBuffer);
  }

  static async appendIf(
    element: HelperElement,
  ): Promise<HelperElement | boolean> {
    const xml = await XmlHelper.getXmlFromArchive(
      element.archive,
      element.file,
    );

    return element.clause !== undefined && !element.clause(xml)
      ? false
      : XmlHelper.append(element);
  }

  static async append(element: HelperElement): Promise<HelperElement> {
    const xml = await XmlHelper.getXmlFromArchive(
      element.archive,
      element.file,
    );

    const newElement = xml.createElement(element.tag);
    for (const attribute in element.attributes) {
      const value = element.attributes[attribute];
      newElement.setAttribute(
        attribute,
        typeof value === 'function' ? value(xml) : value,
      );
    }

    if (element.assert) {
      element.assert(xml);
    }

    const parent = element.parent(xml);
    parent.appendChild(newElement);

    await XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return newElement as unknown as HelperElement;
  }

  static async getNextRelId(rootArchive: JSZip, file: string): Promise<string> {
    const presentationRelsXml = await XmlHelper.getXmlFromArchive(
      rootArchive,
      file,
    );
    const increment = (max: number) => 'rId' + max;
    const relationNodes = presentationRelsXml.documentElement.childNodes;
    const rid = XmlHelper.getMaxId(relationNodes, 'Id', true);

    return increment(rid) + '-created';
  }

  static getMaxId(
    rels: NodeListOf<ChildNode> | HTMLCollectionOf<Element>,
    attribute: string,
    increment?: boolean,
    minId?: number,
  ): number {
    let max = 0;
    for (const i in rels) {
      const rel = rels[i] as Element;
      if (rel.getAttribute !== undefined) {
        const id = Number(
          rel
            .getAttribute(attribute)
            .replace('rId', '')
            .replace('-created', ''),
        );
        max = id > max ? id : max;
      }
    }

    switch (typeof increment) {
      case 'boolean':
        ++max;
        break;
    }

    if (max < minId) {
      return minId;
    }

    return max;
  }

  static async getTargetsFromRelationships(
    archive: JSZip,
    path: string,
    prefix: string,
    suffix?: string | RegExp,
  ): Promise<Target[]> {
    return XmlHelper.getRelationships(
      archive,
      path,
      (element: Element, rels: Target[]) => {
        const target = element.getAttribute('Target');
        if (target.indexOf(prefix) === 0) {
          rels.push({
            file: target,
            rId: element.getAttribute('Id'),
            number: Number(
              target.replace(prefix, '').replace(suffix || '.xml', ''),
            ),
          } as Target);
        }
      },
    );
  }

  static async getTargetsByRelationshipType(
    archive: JSZip,
    path: string,
    type: string,
  ): Promise<Target[]> {
    return XmlHelper.getRelationships(
      archive,
      path,
      (element: Element, rels: Target[]) => {
        const target = element.getAttribute('Type');
        if (target === type) {
          rels.push({
            file: element.getAttribute('Target'),
            rId: element.getAttribute('Id'),
          } as Target);
        }
      },
    );
  }

  static async getRelationships(
    archive: JSZip,
    path: string,
    cb: GetRelationshipsCallback,
  ): Promise<Target[]> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const relationships = xml.getElementsByTagName('Relationship');
    const rels = [];

    Object.keys(relationships)
      .map((key) => relationships[key] as Element)
      .filter((element) => element.getAttribute !== undefined)
      .forEach((element) => cb(element, rels));

    return rels;
  }

  static findByAttribute(
    xml: XMLDocument | Document,
    tagName: string,
    attributeName: string,
    attributeValue: string,
  ): boolean {
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

  static async replaceAttribute(
    archive: JSZip,
    path: string,
    tagName: string,
    attributeName: string,
    attributeValue: string,
    replaceValue: string,
  ): Promise<JSZip> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const elements = xml.getElementsByTagName(tagName);
    for (const i in elements) {
      const element = elements[i];
      if (
        element.getAttribute !== undefined &&
        element.getAttribute(attributeName) === attributeValue
      ) {
        element.setAttribute(attributeName, replaceValue);
      }
    }
    return XmlHelper.writeXmlToArchive(archive, path, xml);
  }

  static async getTargetByRelId(
    archive: JSZip,
    slideNumber: number,
    element: XMLDocument,
    type: string,
  ): Promise<Target> {
    const params = TargetByRelIdMap[type];
    const sourceRid = element
      .getElementsByTagName(params.relRootTag)[0]
      .getAttribute(params.relAttribute);
    const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
    const imageRels = await XmlHelper.getTargetsFromRelationships(
      archive,
      relsPath,
      params.prefix,
      params.expression,
    );
    const target = imageRels.find((rel) => rel.rId === sourceRid);

    return target;
  }

  static async findByElementCreationId(
    archive: JSZip,
    path: string,
    creationId: string,
  ): Promise<XMLDocument> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByCreationId(slideXml, creationId);
  }

  static async findByElementName(
    archive: JSZip,
    path: string,
    name: string,
  ): Promise<XMLDocument> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByName(slideXml, name);
  }

  static findByName(doc: Document, name: string): XMLDocument {
    const names = doc.getElementsByTagName('p:cNvPr');

    for (const i in names) {
      if (names[i].getAttribute && names[i].getAttribute('name') === name) {
        return names[i].parentNode.parentNode as XMLDocument;
      }
    }

    return null;
  }

  static findByCreationId(doc: Document, creationId: string): XMLDocument {
    const creationIds = doc.getElementsByTagName('a16:creationId');

    for (const i in creationIds) {
      if (
        creationIds[i].getAttribute &&
        creationIds[i].getAttribute('id') === creationId
      ) {
        return creationIds[i].parentNode.parentNode.parentNode.parentNode
          .parentNode as XMLDocument;
      }
    }

    return null;
  }

  static findFirstByAttributeValue(
    nodes: NodeListOf<ChildNode> | HTMLCollectionOf<Element>,
    attributeName: string,
    attributeValue: string,
  ): Element {
    for (const i in nodes) {
      const node = <Element>nodes[i];
      if (
        node.getAttribute &&
        node.getAttribute(attributeName) === attributeValue
      ) {
        return node;
      }
    }
    return null;
  }

  static findByAttributeValue(
    nodes: NodeListOf<ChildNode> | HTMLCollectionOf<Element>,
    attributeName: string,
    attributeValue: string,
  ): Element[] {
    const matchingNodes = <Element[]>[];
    for (const i in nodes) {
      const node = <Element>nodes[i];
      if (
        node.getAttribute &&
        node.getAttribute(attributeName) === attributeValue
      ) {
        matchingNodes.push(node);
      }
    }
    return matchingNodes;
  }

  static createContentTypeChild(
    archive: JSZip,
    attributes: OverrideAttribute | DefaultAttribute,
  ): HelperElement {
    return {
      archive,
      file: `[Content_Types].xml`,
      parent: (xml: XMLDocument) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes,
    };
  }

  static createRelationshipChild(
    archive: JSZip,
    targetRelFile: string,
    attributes: RelationshipAttribute,
  ): HelperElement {
    return {
      archive,
      file: targetRelFile,
      parent: (xml: XMLDocument) =>
        xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes,
    };
  }

  static appendSharedString(
    sharedStrings: Document,
    stringValue: string,
  ): number {
    const strings = sharedStrings.getElementsByTagName('sst')[0];
    const newLabel = sharedStrings.createTextNode(stringValue);
    const newText = sharedStrings.createElement('t');
    newText.appendChild(newLabel);

    const newString = sharedStrings.createElement('si');
    newString.appendChild(newText);

    strings.appendChild(newString);

    return strings.getElementsByTagName('si').length - 1;
  }

  static insertAfter(newNode: Node, referenceNode: Element): Node {
    return referenceNode.parentNode.insertBefore(
      newNode,
      referenceNode.nextSibling,
    );
  }

  static sliceCollection(
    collection: HTMLCollectionOf<Element>,
    length: number,
    from?: number,
  ): void {
    if (from !== undefined) {
      for (let i = from; i < length; i++) {
        const toRemove = collection[i];
        toRemove.parentNode.removeChild(toRemove);
      }
    } else {
      for (let i = collection.length; i > length; i--) {
        const toRemove = collection[i - 1];
        toRemove.parentNode.removeChild(toRemove);
      }
    }
  }

  static dump(element: XMLDocument | Element): void {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(element);
    const p = new XmlPrettyPrint(xmlBuffer);
    p.dump();
  }
}
