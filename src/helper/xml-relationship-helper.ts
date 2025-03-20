import { XmlDocument, XmlElement } from '../types/xml-types';
import { Target } from '../types/types';
import IArchive from '../interfaces/iarchive';
import { XmlHelper } from './xml-helper';
import { last, vd } from './general-helper';
import { ElementSubtype } from '../enums/element-type';
import { FileHelper } from './file-helper';
import { randomBytes } from 'crypto';

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
    const fileProxy = await this.archive;
    this.xml = await XmlHelper.getXmlFromArchive(
      fileProxy,
      this.path + this.file,
    );

    await this.readTargets();

    if (prefix) {
      return this.getTargetsByPrefix(prefix);
    }

    return this;
  }

  setXml(xml) {
    this.xml = xml;
    return this;
  }

  getTargetsByPrefix(prefix: string | string[]): Target[] {
    const prefixes = typeof prefix === 'string' ? [prefix] : prefix;

    const targets = [];
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
    const targets = [];
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
          xmlTarget.setAttribute('Target', targetFile + targetSuffix);

          await XmlHelper.appendImageExtensionToContentType(
            this.archive,
            target.filenameExt,
          );
        }
      }
    }
  }

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
        target.element.setAttribute('Target', newTarget);
      },
      updateId: (newId: string) => {
        target.element.setAttribute('Id', newId);
      },
    };

    if (prefix) {
      const subtype = last(prefix.split('/')) as ElementSubtype;

      if (
        matchByPrefix &&
        !XmlRelationshipHelper.targetMatchesRelationship(
          relType,
          subtype,
          file,
          prefix,
        )
      ) {
        return;
      }
      return this.extendTarget(prefix, subtype, target);
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
        target.element.setAttribute('Target', `${prefix}${newIndex}.xml`);
      },
    };
  }

  static targetMatchesRelationship(relType, subtype, target, prefix) {
    if (relType === 'package') return true;

    return relType === subtype && target.indexOf(prefix) === 0;
  }

  static async getSlideLayoutNumber(sourceArchive, slideId: number) {
    const slideToLayouts = await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `slide${slideId}.xml.rels`,
      `ppt/slides/_rels`,
      '../slideLayouts/slideLayout',
    );

    return slideToLayouts[0].number;
  }

  static async getSlideMasterNumber(sourceArchive, slideLayoutId: number) {
    const layoutToMaster = (await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `slideLayout${slideLayoutId}.xml.rels`,
      `ppt/slideLayouts/_rels`,
      '../slideMasters/slideMaster',
    )) as Target[];

    return layoutToMaster[0].number;
  }
}
