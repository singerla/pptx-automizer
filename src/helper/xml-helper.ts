import { Node, XMLSerializer } from '@xmldom/xmldom';
import {
  DefaultAttribute,
  HelperElement,
  ModifyXmlCallback,
  OverrideAttribute,
  RelationshipAttribute,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { TargetByRelIdMap } from '../constants/constants';
import { XmlPrettyPrint } from './xml-pretty-print';
import { GetRelationshipsCallback, Target } from '../types/types';
import { log } from './general-helper';
import { contentTracker } from './content-tracker';
import IArchive from '../interfaces/iarchive';
import {
  ContentTypeExtension,
  ContentTypeMap,
} from '../enums/content-type-map';

export class XmlHelper {
  static async modifyXmlInArchive(
    archive: IArchive,
    file: string,
    callbacks: ModifyXmlCallback[],
  ): Promise<void> {
    const fileProxy = await archive;
    const xml = await XmlHelper.getXmlFromArchive(fileProxy, file);

    let i = 0;
    for (const callback of callbacks) {
      await callback(xml, i++, fileProxy);
    }

    XmlHelper.writeXmlToArchive(await archive, file, xml);
  }

  static async getXmlFromArchive(
    archive: IArchive,
    file: string,
  ): Promise<XmlDocument> {
    return archive.readXml(file);
  }

  static writeXmlToArchive(
    archive: IArchive,
    file: string,
    xml: XmlDocument,
  ): void {
    archive.writeXml(file, xml);
  }

  static async appendIf(element: HelperElement): Promise<XmlElement | boolean> {
    const xml = await XmlHelper.getXmlFromArchive(
      element.archive,
      element.file,
    );

    return element.clause !== undefined && !element.clause(xml)
      ? false
      : XmlHelper.append(element);
  }

  static async append(element: HelperElement): Promise<XmlElement> {
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

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return newElement as XmlElement;
  }

  static async removeIf(element: HelperElement): Promise<XmlElement[]> {
    const xml = await XmlHelper.getXmlFromArchive(
      element.archive,
      element.file,
    );

    const collection = xml.getElementsByTagName(element.tag);
    const toRemove: XmlElement[] = [];
    XmlHelper.modifyCollection(collection, (item: XmlElement, index) => {
      if (element.clause(xml, item)) {
        toRemove.push(item);
      }
    });

    toRemove.forEach((item) => {
      XmlHelper.remove(item);
    });

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml);

    return toRemove;
  }

  static async getNextRelId(
    rootArchive: IArchive,
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
    rels: NodeListOf<ChildNode> | HTMLCollectionOf<XmlElement>,
    attribute: string,
    increment?: boolean,
    minId?: number,
  ): number {
    let max = 0;
    for (const i in rels) {
      const rel = rels[i] as XmlElement;
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

  static async getRelationshipTargetsByPrefix(
    archive: IArchive,
    path: string,
    prefix: string | string[],
  ): Promise<Target[]> {
    const prefixes = typeof prefix === 'string' ? [prefix] : prefix;
    return XmlHelper.getRelationshipItems(
      archive,
      path,
      (element: XmlElement, targets: Target[]) => {
        prefixes.forEach((prefix) => {
          const target = XmlHelper.parseRelationTarget(element, prefix);
          if (target.prefix) {
            targets.push(target);
          }
        });
      },
    );
  }

  static parseRelationTarget(element: XmlElement, prefix?: string): Target {
    const type = element.getAttribute('Type');
    const file = element.getAttribute('Target');

    const last = (arr: string[]): string => arr[arr.length - 1];
    const filename = last(file.split('/'));
    const subtype = last(prefix.split('/'));

    const relType = last(type.split('/'));
    const rId = element.getAttribute('Id');
    const filenameExt = last(filename.split('.'));
    const filenameMatch = filename
      .replace('.' + filenameExt, '')
      .match(/^(.+?)(\d+)*$/);
    const filenameBase =
      filenameMatch && filenameMatch[1] ? filenameMatch[1] : filename;
    const number =
      filenameMatch && filenameMatch[2] ? Number(filenameMatch[2]) : 0;

    const target = <Target>{
      rId,
      type,
      file,
      filename,
      filenameBase,
      number,
      subtype,
      relType,
      element,
    };

    if (
      prefix &&
      XmlHelper.targetMatchesRelationship(relType, subtype, file, prefix)
    ) {
      return {
        ...target,
        prefix,
      } as Target;
    }

    if (prefix && prefix.indexOf('../') === 0) {
      // Try again with absolute path instead of relative
      return XmlHelper.parseRelationTarget(
        element,
        prefix.replace('../', '/ppt/'),
      );
    }

    return target;
  }

  static targetMatchesRelationship(
    relType: string,
    subtype: string,
    file: string,
    prefix: string,
  ) {
    if (relType === 'package') return true;

    // pptgenjs uses absolute paths in "Target" attributes
    if (file.indexOf('/ppt/') === 0) {
      file = file.replace('/ppt/', '../');
    }

    return relType === subtype && file.indexOf(prefix) === 0;
  }

  static async getTargetsByRelationshipType(
    archive: IArchive,
    path: string,
    type: string,
  ): Promise<Target[]> {
    return await XmlHelper.getRelationshipItems(
      archive,
      path,
      (element: XmlElement, rels: Target[]) => {
        const target = element.getAttribute('Type');
        if (target === type) {
          rels.push({
            file: element.getAttribute('Target'),
            rId: element.getAttribute('Id'),
            element: element,
          } as Target);
        }
      },
    );
  }

  static async getRelationshipItems(
    archive: IArchive,
    path: string,
    cb: GetRelationshipsCallback,
    tag?: string,
  ): Promise<Target[]> {
    tag = tag || 'Relationship';

    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const relationshipItems = xml.getElementsByTagName(tag);

    const rels = [];

    for (const i in relationshipItems) {
      if (relationshipItems[i].getAttribute) {
        cb(relationshipItems[i], rels);
      }
    }

    return rels;
  }

  static findByAttribute(
    xml: XmlDocument | Document,
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
    archive: IArchive,
    path: string,
    tagName: string,
    attributeName: string,
    attributeValue: string,
    replaceValue: string,
    replaceAttributeName?: string,
  ): Promise<void> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const elements = xml.getElementsByTagName(tagName);
    for (const i in elements) {
      const element = elements[i];
      if (
        element.getAttribute !== undefined &&
        element.getAttribute(attributeName) === attributeValue
      ) {
        element.setAttribute(
          replaceAttributeName || attributeName,
          replaceValue,
        );
      }

      if (element.getAttribute !== undefined) {
        contentTracker.trackRelation(path, {
          Id: element.getAttribute('Id'),
          Target: element.getAttribute('Target'),
          Type: element.getAttribute('Type'),
        });
      }
    }
    XmlHelper.writeXmlToArchive(archive, path, xml);
  }

  static async getTargetByRelId(
    archive: IArchive,
    relsPath: string,
    element: XmlElement,
    type: string,
  ): Promise<Target> {
    const params = TargetByRelIdMap[type];

    // For elements that need to search all instances (like hyperlinks)
    if (params.findAll) {
      // Find all hyperlink elements
      const hyperlinks = element.getElementsByTagName(params.relRootTag);
      if (hyperlinks.length > 0) {
        // Use the first hyperlink found
        const sourceRid = hyperlinks[0].getAttribute(params.relAttribute);

        // Get all relationships
        const allRels = await XmlHelper.getRelationshipItems(
          archive,
          relsPath,
          (element: XmlElement, rels: Target[]) => {
            rels.push({
              rId: element.getAttribute('Id'),
              type: element.getAttribute('Type'),
              file: element.getAttribute('Target'),
              filename: element.getAttribute('Target'),
              element: element,
              isExternal: element.getAttribute('TargetMode') === 'External',
            } as Target);
          },
        );

        // Find the matching relationship
        const target = allRels.find((rel) => rel.rId === sourceRid);
        return target;
      }
    } else {
      // Standard behavior for other element types
      const sourceRid = element
        .getElementsByTagName(params.relRootTag)
        .item(0)
        ?.getAttribute(params.relAttribute);

      if (!sourceRid) {
        throw 'No sourceRid for ' + params.relRootTag;
      }

      const shapeRels = await XmlHelper.getRelationshipTargetsByPrefix(
        archive,
        relsPath,
        params.prefix,
      );

      const target = shapeRels.find((rel) => {
        return rel.rId === sourceRid;
      });

      return target;
    }
  }

  // Determine whether a given string is a creationId or a shape name
  // Example creationId: '{EFC74B4C-D832-409B-9CF4-73C1EFF132D8}'
  static isElementCreationId(selector: string) {
    return selector.indexOf('{') === 0 && selector.split('-').length === 5;
  }

  static async findByElementCreationId(
    archive: IArchive,
    path: string,
    creationId: string,
  ): Promise<XmlElement> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByCreationId(slideXml, creationId);
  }

  static async findByElementName(
    archive: IArchive,
    path: string,
    name: string,
  ): Promise<XmlElement> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

    return XmlHelper.findByName(slideXml, name);
  }

  static findByName(doc: Document, name: string): XmlElement {
    const names = doc.getElementsByTagName('p:cNvPr');

    for (const i in names) {
      if (names[i].getAttribute && names[i].getAttribute('name') === name) {
        return names[i].parentNode.parentNode as XmlElement;
      }
    }

    return null;
  }

  static findByCreationId(doc: Document, creationId: string): XmlElement {
    const creationIds = doc.getElementsByTagName('a16:creationId');

    for (const i in creationIds) {
      if (
        creationIds[i].getAttribute &&
        creationIds[i].getAttribute('id') === creationId
      ) {
        return creationIds[i].parentNode.parentNode.parentNode.parentNode
          .parentNode as XmlElement;
      }
    }

    return null;
  }

  static findFirstByAttributeValue(
    nodes: NodeListOf<ChildNode> | HTMLCollectionOf<XmlElement>,
    attributeName: string,
    attributeValue: string,
  ): XmlElement {
    for (const i in nodes) {
      const node = <XmlElement>nodes[i];
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
    nodes: NodeListOf<ChildNode> | HTMLCollectionOf<XmlElement>,
    attributeName: string,
    attributeValue: string,
  ): XmlElement[] {
    const matchingNodes = <XmlElement[]>[];
    for (const i in nodes) {
      const node = <XmlElement>nodes[i];
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
    archive: IArchive,
    attributes: OverrideAttribute | DefaultAttribute,
  ): HelperElement {
    return {
      archive,
      file: `[Content_Types].xml`,
      parent: (xml: XmlDocument) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes,
    };
  }

  static createRelationshipChild(
    archive: IArchive,
    targetRelFile: string,
    attributes: RelationshipAttribute,
  ): HelperElement {
    contentTracker.trackRelation(targetRelFile, attributes);

    return {
      archive,
      file: targetRelFile,
      parent: (xml: XmlDocument) =>
        xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes,
    };
  }

  static appendImageExtensionToContentType(
    targetArchive: IArchive,
    extension: ContentTypeExtension,
  ): Promise<XmlElement | boolean> {
    const contentType = ContentTypeMap[extension]
      ? ContentTypeMap[extension]
      : 'image/' + extension;

    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(targetArchive, {
        Extension: extension,
        ContentType: contentType,
      }),
      tag: 'Default',
      clause: (xml: XmlDocument) =>
        !XmlHelper.findByAttribute(xml, 'Default', 'Extension', extension),
    });
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

  static insertAfter(
    newNode: XmlElement,
    referenceNode: XmlElement,
  ): XmlElement {
    return referenceNode.parentNode.insertBefore(
      newNode,
      referenceNode.nextSibling,
    );
  }

  static sliceCollection(
    collection: HTMLCollectionOf<XmlElement>,
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

  static getClosestParent(tag: string, element: XmlElement): XmlElement {
    if (element.parentNode) {
      if (element.parentNode.nodeName === tag) {
        return element.parentNode as XmlElement;
      }
      return XmlHelper.getClosestParent(tag, element.parentNode as XmlElement);
    }
  }

  static remove(toRemove: XmlElement): void {
    if (toRemove?.parentNode) {
      toRemove.parentNode.removeChild(toRemove);
    }
  }

  static moveChild(childToMove: XmlElement, insertBefore?: XmlElement): void {
    const parent = childToMove.parentNode;
    parent.insertBefore(childToMove, insertBefore);
  }

  static appendClone(childToClone: XmlElement, parent: XmlElement): XmlElement {
    const clone = childToClone.cloneNode(true) as XmlElement;
    parent.appendChild(clone);
    return clone;
  }

  static sortCollection(
    collection: HTMLCollectionOf<XmlElement>,
    order: number[],
    callback?: ModifyXmlCallback,
  ): void {
    if (collection.length === 0) {
      return;
    }
    const parent = collection[0].parentNode;
    order.forEach((index, i) => {
      if (!collection[index]) {
        log('sortCollection index not found' + index, 1);
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
    collection: HTMLCollectionOf<XmlElement>,
    callback: ModifyXmlCallback,
  ): void {
    for (let i = 0; i < collection.length; i++) {
      const item = collection[i];
      callback(item, i);
    }
  }

  static async modifyCollectionAsync(
    collection: HTMLCollectionOf<XmlElement>,
    callback: ModifyXmlCallback,
  ): Promise<void> {
    for (let i = 0; i < collection.length; i++) {
      const item = collection[i];
      await callback(item, i);
    }
  }

  static dump(element: XMLDocument | Element | Node): void {
    const s = new XMLSerializer();
    const xmlBuffer = s.serializeToString(<Node>element);
    const p = new XmlPrettyPrint(xmlBuffer);
    p.dump();
  }
}
