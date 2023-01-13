import JSZip from 'jszip';

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { FileHelper } from './file-helper';
import {
  DefaultAttribute,
  HelperElement,
  OverrideAttribute,
  RelationshipAttribute,
  ModifyXmlCallback,
} from '../types/xml-types';
import { TargetByRelIdMap } from '../constants/constants';
import { XmlPrettyPrint } from './xml-pretty-print';
import { GetRelationshipsCallback, Target } from '../types/types';
import _ from 'lodash';
import { vd } from './general-helper';
import { contentTracker, ContentTracker } from './content-tracker';
import { FileProxy } from './file-proxy';

export class XmlHelper {
  static async modifyXmlInArchive(
    archive: FileProxy,
    file: string,
    callbacks: ModifyXmlCallback[],
  ): Promise<FileProxy> {
    const fileProxy = await archive;
    const xml = await XmlHelper.getXmlFromArchive(fileProxy, file);

    let i = 0;
    for (const callback of callbacks) {
      await callback(xml, i++, fileProxy);
    }

    return await XmlHelper.writeXmlToArchive(await archive, file, xml);
  }

  static async getXmlFromArchive(
    archive: FileProxy,
    file: string,
  ): Promise<XMLDocument> {
    const xmlDocument = (await archive.read(file, 'string')) as string;
    const dom = new DOMParser();
    return dom.parseFromString(xmlDocument);
  }

  static async writeXmlToArchive(
    archive: FileProxy,
    file: string,
    xml: XMLDocument,
  ): Promise<FileProxy> {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(xml);

    return archive.write(file, xmlBuffer);
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
      const setValue = typeof value === 'function' ? value(xml) : value;

      newElement.setAttribute(attribute, setValue);
    }

    contentTracker.trackRelation(
      element.file,
      element.attributes as RelationshipAttribute,
    );

    if (element.assert) {
      element.assert(xml);
    }

    const parent = element.parent(xml);
    parent.appendChild(newElement);

    await XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return newElement as unknown as HelperElement;
  }

  static async removeIf(element: HelperElement): Promise<Element[]> {
    const xml = await XmlHelper.getXmlFromArchive(
      element.archive,
      element.file,
    );

    const collection = xml.getElementsByTagName(element.tag);
    const toRemove: Element[] = [];
    XmlHelper.modifyCollection(collection, (item: Element, index) => {
      if (element.clause(xml, item)) {
        toRemove.push(item);
      }
    });

    toRemove.forEach((item) => {
      XmlHelper.remove(item);
    });

    await XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return toRemove;
  }

  static async getNextRelId(
    rootArchive: FileProxy,
    file: string,
  ): Promise<string> {
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
    archive: FileProxy,
    path: string,
    prefix: string | string[],
  ): Promise<Target[]> {
    const prefixes = typeof prefix === 'string' ? [prefix] : prefix;

    return XmlHelper.getRelationships(
      archive,
      path,
      (element: Element, targets: Target[]) => {
        prefixes.forEach((prefix) => {
          XmlHelper.pushRelTargets(element, prefix, targets);
        });
      },
    );
  }

  static pushRelTargets(element: Element, prefix: string, targets: Target[]) {
    const type = element.getAttribute('Type');
    const file = element.getAttribute('Target');
    const rId = element.getAttribute('Id');

    const subtype = _.last(prefix.split('/'));
    const relType = _.last(type.split('/'));
    const filename = _.last(file.split('/'));
    const filenameExt = _.last(filename.split('.'));
    const filenameMatch = filename
      .replace('.' + filenameExt, '')
      .match(/^(.+?)(\d+)*$/);

    const number =
      filenameMatch && filenameMatch[2] ? Number(filenameMatch[2]) : 0;
    const filenameBase =
      filenameMatch && filenameMatch[1] ? filenameMatch[1] : filename;

    if (XmlHelper.targetMatchesRelationship(relType, subtype, file, prefix)) {
      targets.push({
        file,
        rId,
        number,
        type,
        subtype,
        prefix,
        filename,
        filenameExt,
        filenameBase,
      } as Target);
    }
  }

  static targetMatchesRelationship(relType, subtype, target, prefix) {
    if (relType === 'package') return true;

    return relType === subtype && target.indexOf(prefix) === 0;
  }

  static async getTargetsByRelationshipType(
    archive: FileProxy,
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
    archive: FileProxy,
    path: string,
    cb: GetRelationshipsCallback,
  ): Promise<Target[]> {
    return this.getRelationshipItems(archive, path, 'Relationship', cb);
  }

  static async getRelationshipItems(
    archive: FileProxy,
    path: string,
    tag: string,
    cb: GetRelationshipsCallback,
  ): Promise<Target[]> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const relationshipItems = xml.getElementsByTagName(tag);
    const rels = [];

    Object.keys(relationshipItems)
      .map((key) => relationshipItems[key] as Element)
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
    archive: FileProxy,
    path: string,
    tagName: string,
    attributeName: string,
    attributeValue: string,
    replaceValue: string,
  ): Promise<FileProxy> {
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

      if (element.getAttribute !== undefined) {
        contentTracker.trackRelation(path, {
          Id: element.getAttribute('Id'),
          Target: element.getAttribute('Target'),
          Type: element.getAttribute('Type'),
        });
      }
    }
    return XmlHelper.writeXmlToArchive(archive, path, xml);
  }

  static async getTargetByRelId(
    archive: FileProxy,
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
    );
    const target = imageRels.find((rel) => rel.rId === sourceRid);

    return target;
  }

  static async findByElementCreationId(
    archive: FileProxy,
    path: string,
    creationId: string,
  ): Promise<XMLDocument> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByCreationId(slideXml, creationId);
  }

  static async findByElementName(
    archive: FileProxy,
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
    archive: FileProxy,
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
    archive: FileProxy,
    targetRelFile: string,
    attributes: RelationshipAttribute,
  ): HelperElement {
    contentTracker.trackRelation(targetRelFile, attributes);

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
        XmlHelper.remove(collection[i]);
      }
    } else {
      for (let i = collection.length; i > length; i--) {
        XmlHelper.remove(collection[i - 1]);
      }
    }
  }

  static remove(toRemove: Element): void {
    toRemove.parentNode.removeChild(toRemove);
  }

  static sortCollection(
    collection: HTMLCollectionOf<Element>,
    order: number[],
    callback?: ModifyXmlCallback,
  ): void {
    if (collection.length === 0) {
      return;
    }
    const parent = collection[0].parentNode;
    order.forEach((index, i) => {
      if (!collection[index]) {
        vd('sortCollection index not found' + index);
        return;
      }

      const item = collection[index];
      if (callback) {
        callback(item, i);
      }
      parent.appendChild(item);
    });
  }

  static modifyCollection(
    collection: HTMLCollectionOf<Element>,
    callback: ModifyXmlCallback,
  ): void {
    for (let i = 0; i < collection.length; i++) {
      const item = collection[i];
      callback(item, i);
    }
  }

  static dump(element: XMLDocument | Element): void {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(element);
    const p = new XmlPrettyPrint(xmlBuffer);
    p.dump();
  }
}
