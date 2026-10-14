import {
  HelperElement,
  RelationshipAttribute,
  XmlDocument,
  XmlElement,
  XmlElementCollection,
  XmlNodeCollection,
} from '../types/xml-types';
import { GetRelationshipsCallback, Target } from '../types/types';
import IArchive from '../interfaces/iarchive';
import { XmlHelper } from './xml-helper';
import { last } from './general-helper';
import { ElementSubtype } from '../enums/element-type';
import { FileHelper } from './file-helper';
import { randomBytes } from 'crypto';
import { TargetByRelIdMap } from '../constants/constants';
import { ElementNotFoundError } from '../errors';

/**
 * Relationship subtypes that are spelled differently in the target prefix
 * and in the relationship type: a `../diagrams/data` prefix refers to a
 * `.../relationships/diagramData` relationship.
 */
const PREFIX_SUBTYPE_MAP: Record<string, string> = {
  data: 'diagramData',
};

/**
 * Relationship handling for `_rels` parts (ROADMAP: #55).
 *
 * Instance methods work on one parsed rels document (`initialize` /
 * `setXml`, then `getTargetsByPrefix`, `getTargetsByType`, …). The static
 * methods are the archive-level helpers that used to live on `XmlHelper`;
 * the `XmlHelper` names remain as deprecated delegates.
 */

export class XmlRelationshipHelper {
  archive: IArchive;
  file: string;
  path: string;
  tag: string;
  xml: XmlDocument;
  xmlTargets: XmlElement[] = [];
  targets: Target[] = [];

  constructor(xml?: XmlDocument, tag?: string) {
    if (xml) {
      this.setXml(xml);
    }
    this.tag = tag || 'Relationship';
    return this;
  }

  async initialize(
    archive: IArchive,
    file: string,
    path: string,
    prefix?: string,
  ) {
    this.archive = archive;
    this.file = file;
    this.path = path + '/';
    this.xml = await XmlHelper.getXmlFromArchive(
      this.archive,
      this.path + this.file,
    );

    await this.readTargets();

    if (prefix) {
      return this.getTargetsByPrefix(prefix);
    }

    return this;
  }

  setXml(xml: XmlDocument) {
    this.xml = xml;
    return this;
  }

  getTargetsByPrefix(prefix: string | string[]): Target[] {
    const prefixes = typeof prefix === 'string' ? [prefix] : prefix;

    const targets: Target[] = [];
    this.xmlTargets.forEach((xmlTarget) => {
      prefixes.forEach((prefix) => {
        const target = XmlRelationshipHelper.parseRelationTarget(
          xmlTarget,
          prefix,
          true,
        );
        if (target?.prefix) {
          targets.push(target);
        }
      });
    });

    return targets;
  }

  getTargetsByType(type: string): Target[] {
    const targets: Target[] = [];
    this.xmlTargets.forEach((xmlTarget) => {
      const target = XmlRelationshipHelper.parseRelationTarget(xmlTarget);
      if (target?.type === type) {
        targets.push(target);
      }
    });
    return targets;
  }

  getTargetByRelId(findRid: string): Target | null {
    const matchedTarget = this.xmlTargets.find(
      (xmlTarget) => xmlTarget.getAttribute('Id') === findRid,
    );

    if (matchedTarget) {
      return XmlRelationshipHelper.parseRelationTarget(matchedTarget);
    }
  }

  readTargets(): this {
    if (this.xmlTargets.length) {
      return this;
    }

    const relationshipItems = this.xml.getElementsByTagName(this.tag);

    for (const i in relationshipItems) {
      if (
        relationshipItems[i] &&
        relationshipItems[i].getAttribute !== undefined
      ) {
        this.xmlTargets.push(relationshipItems[i]);
      }
    }

    return this;
  }

  /**
   * This will copy all unhandled related contents into
   * the target archive.
   *
   * Pptx messages on opening a corrupted file are most likely
   * caused by broken relations and this is going to prevent
   * files from being missed.
   *
   * @param sourceArchive
   * @param check
   * @param assert
   */
  async assertRelatedContent(
    sourceArchive: IArchive,
    check?: boolean,
    assert?: boolean,
  ) {
    for (const xmlTarget of this.xmlTargets) {
      const targetFile = xmlTarget.getAttribute('Target');
      const targetMode = xmlTarget.getAttribute('TargetMode');
      const targetPath = targetFile.replace('../', 'ppt/');

      if (
        targetMode !== 'External' &&
        this.archive.fileExists(targetPath) === false
      ) {
        // ToDo: There are falsy errors on files that have already been
        //       copied with another target name.
        // if (check) {
        //   if (typeof sourceArchive.filename === 'string') {
        //     console.error(
        //       'Related content from ' +
        //         sourceArchive.filename +
        //         ' not found: ' +
        //         targetFile,
        //     );
        //   } else {
        //     console.error('Related content not found: ' + targetFile);
        //   }
        // }

        if (assert) {
          const target = XmlRelationshipHelper.parseRelationTarget(xmlTarget);
          const buf = randomBytes(5).toString('hex');
          const targetSuffix = '-' + buf + '.' + target.filenameExt;
          await FileHelper.zipCopy(
            sourceArchive,
            targetPath,
            this.archive,
            targetPath + targetSuffix,
          );
          xmlTarget.setAttribute(
            'Target',
            XmlHelper.sanitizeAttr(targetFile + targetSuffix),
          );

          await XmlHelper.appendImageExtensionToContentType(
            this.archive,
            target.filenameExt,
          );
        }
      }
    }
  }

  /**
   * Parses one `<Relationship>` element into a `Target`.
   *
   * With a `prefix` (e.g. `../charts/chart`), the target is matched against
   * it: the relationship type must end in the prefix's last path segment
   * (`chart`; `data` is mapped to `diagramData`) and the `Target` attribute
   * must start with the prefix. Absolute `/ppt/...` targets, as written by
   * pptxgenjs, are matched as if they were relative. A matching target gets
   * `prefix`, `subtype` and `updateTargetIndex()`; a non-matching one is
   * returned without `prefix` — or as `undefined` if `matchByPrefix` is set.
   *
   * `number` is the trailing number of the target file name (`chart12.xml`
   * → 12, `0` if there is none), `filenameBase` the part before it.
   */
  static parseRelationTarget(
    element: XmlElement,
    prefix?: string,
    matchByPrefix?: boolean,
  ): Target | undefined {
    if (!element || element.getAttribute === undefined) {
      return;
    }

    const type = element.getAttribute('Type');
    const file = element.getAttribute('Target');
    const rId = element.getAttribute('Id');

    const filename = last(file.split('/'));
    const relType = last(type.split('/'));
    const filenameExt = last(filename.split('.'));
    const filenameMatch = filename
      .replace('.' + filenameExt, '')
      .match(/^(.+?)(\d+)*$/);

    const number =
      filenameMatch && filenameMatch[2] ? Number(filenameMatch[2]) : 0;
    const filenameBase =
      filenameMatch && filenameMatch[1] ? filenameMatch[1] : filename;

    const target = <Target>{
      rId,
      type,
      file,
      filename,
      relType,
      element,
      filenameExt,
      filenameMatch,
      number,
      filenameBase,
      getTargetValue: () => target.element.getAttribute('Target'),
      updateTargetValue: (newTarget: string) => {
        target.element.setAttribute('Target', XmlHelper.sanitizeAttr(newTarget));
      },
      updateId: (newId: string) => {
        target.element.setAttribute('Id', newId);
      },
    };

    if (prefix) {
      const prefixSubtype = last(prefix.split('/'));
      const subtype = (PREFIX_SUBTYPE_MAP[prefixSubtype] ||
        prefixSubtype) as ElementSubtype;

      if (
        XmlRelationshipHelper.targetMatchesRelationship(
          relType,
          subtype,
          file,
          prefix,
        )
      ) {
        return XmlRelationshipHelper.extendTarget(prefix, subtype, target);
      }

      if (matchByPrefix) {
        return;
      }

      return { ...target, subtype };
    }

    return target;
  }

  static extendTarget(
    prefix: string,
    subtype: ElementSubtype,
    target: Target,
  ): Target {
    return {
      ...target,
      prefix,
      subtype,
      updateTargetIndex: (newIndex: number) => {
        target.element.setAttribute(
          'Target',
          XmlHelper.sanitizeAttr(`${prefix}${newIndex}.xml`),
        );
      },
    };
  }

  /**
   * `true` if a relationship of `relType` pointing at `file` is the kind of
   * part addressed by `prefix` (see `parseRelationTarget`). Embedded
   * packages (`relType === 'package'`) always match.
   */
  static targetMatchesRelationship(
    relType: string,
    subtype: string,
    file: string,
    prefix: string,
  ): boolean {
    if (relType === 'package') return true;

    // pptxgenjs uses absolute paths in "Target" attributes
    if (file.indexOf('/ppt/') === 0) {
      file = file.replace('/ppt/', '../');
    }

    return relType === subtype && file.indexOf(prefix) === 0;
  }

  /**
   * Iterates the `<Relationship>` elements of a rels part in `archive` and
   * lets `cb` push `Target`s into the returned list.
   */
  static async getRelationshipItems(
    archive: IArchive,
    path: string,
    cb: GetRelationshipsCallback,
    tag?: string,
  ): Promise<Target[]> {
    tag = tag || 'Relationship';

    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    const relationshipItems = xml.getElementsByTagName(tag);

    const rels: Target[] = [];

    for (const i in relationshipItems) {
      if (relationshipItems[i].getAttribute) {
        cb(relationshipItems[i], rels);
      }
    }

    return rels;
  }

  /**
   * All targets of a rels part in `archive` matching one of the given
   * prefixes (see `parseRelationTarget`).
   */
  static async getRelationshipTargetsByPrefix(
    archive: IArchive,
    path: string,
    prefix: string | string[],
  ): Promise<Target[]> {
    const prefixes = typeof prefix === 'string' ? [prefix] : prefix;
    return XmlRelationshipHelper.getRelationshipItems(
      archive,
      path,
      (element: XmlElement, targets: Target[]) => {
        prefixes.forEach((prefix) => {
          const target = XmlRelationshipHelper.parseRelationTarget(
            element,
            prefix,
            true,
          );
          if (target) {
            targets.push(target);
          }
        });
      },
    );
  }

  /**
   * All targets of a rels part in `archive` with the given relationship
   * `Type` URI.
   */
  static async getTargetsByRelationshipType(
    archive: IArchive,
    path: string,
    type: string,
  ): Promise<Target[]> {
    return XmlRelationshipHelper.getRelationshipItems(
      archive,
      path,
      (element: XmlElement, rels: Target[]) => {
        if (element.getAttribute('Type') === type) {
          rels.push(XmlRelationshipHelper.parseRelationTarget(element));
        }
      },
    );
  }

  /**
   * Resolves the relationship target a shape element refers to: reads the
   * element's `r:id`-style attribute for the given `TargetByRelIdMap` key
   * (`chart`, `image`, `hyperlink`, …) and looks it up in `relsPath`.
   */
  static async resolveElementTarget(
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
        const allRels = await XmlRelationshipHelper.getRelationshipItems(
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
        return allRels.find((rel) => rel.rId === sourceRid);
      }
    } else {
      // Standard behavior for other element types
      const sourceRid = element
        .getElementsByTagName(params.relRootTag)
        .item(0)
        ?.getAttribute(params.relAttribute);

      if (!sourceRid) {
        throw new ElementNotFoundError('No sourceRid for ' + params.relRootTag, {
          selector: params.relRootTag,
        });
      }

      const shapeRels =
        await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
          archive,
          relsPath,
          params.prefix,
        );

      return shapeRels.find((rel) => rel.rId === sourceRid);
    }
  }

  /**
   * Next free relationship id of a rels part, suffixed with `-created`
   * (`rId<max+1>-created`). The suffix marks relationships created by
   * automizer; cleanup relies on it — do not change the format.
   */
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
    const rid = XmlRelationshipHelper.getMaxId(relationNodes, 'Id', true);

    return increment(rid) + '-created';
  }

  /**
   * Highest numeric `rId` among `rels` (ignoring the `-created` suffix),
   * optionally incremented by one and raised to `minId`.
   */
  static getMaxId(
    rels: XmlNodeCollection | XmlElementCollection,
    attribute: string,
    increment?: boolean,
    minId?: number,
  ): number {
    let max = 0;
    for (const i in rels) {
      const rel = rels[i] as XmlElement;

      if (rel.getAttribute !== undefined && rel.getAttribute(attribute)) {
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

  /**
   * `HelperElement` for `XmlHelper.append` that adds a `<Relationship>` to
   * `targetRelFile`; the relation is recorded with the archive's content
   * tracker.
   */
  static createRelationshipChild(
    archive: IArchive,
    targetRelFile: string,
    attributes: RelationshipAttribute,
  ): HelperElement {
    archive.contentTracker?.trackRelation(targetRelFile, attributes);

    return {
      archive,
      file: targetRelFile,
      // `Relationships` is the document element of a .rels part.
      parent: (xml: XmlDocument) => xml.documentElement as XmlElement,
      tag: 'Relationship',
      attributes,
    };
  }

  static async getSlideLayoutNumber(sourceArchive: IArchive, slideId: number) {
    const slideToLayouts = (await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `slide${slideId}.xml.rels`,
      `ppt/slides/_rels`,
      '../slideLayouts/slideLayout',
    )) as Target[];
    return slideToLayouts[0].number;
  }

  static async getSlideMasterNumber(
    sourceArchive: IArchive,
    slideLayoutId: number,
  ) {
    const layoutToMaster = (await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `slideLayout${slideLayoutId}.xml.rels`,
      `ppt/slideLayouts/_rels`,
      '../slideMasters/slideMaster',
    )) as Target[];

    return layoutToMaster[0].number;
  }
}
