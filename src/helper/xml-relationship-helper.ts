import { XmlDocument, XmlElement } from '../types/xml-types';
import { Target } from '../types/types';
import IArchive from '../interfaces/iarchive';
import { XmlHelper } from './xml-helper';
import { last, vd } from './general-helper';
import { ElementSubtype } from '../enums/element-type';

export class XmlRelationshipHelper {
  archive: IArchive;
  file: string;
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

  async initialize(archive: IArchive, file: string) {
    this.archive = archive;
    this.file = file;
    const fileProxy = await this.archive;
    this.xml = await XmlHelper.getXmlFromArchive(fileProxy, this.file);

    await this.readTargets();

    return this;
  }

  async writeArchive() {
    await XmlHelper.writeXmlToArchive(this.archive, this.file, this.xml);
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
    const slideToLayoutHelper = await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `ppt/slides/_rels/slide${slideId}.xml.rels`,
    );
    const slideToLayout = slideToLayoutHelper.getTargetsByPrefix(
      '../slideLayouts/slideLayout',
    );
    return slideToLayout[0].number;
  }

  static async getSlideMasterNumber(sourceArchive, slideLayoutId: number) {
    const layoutToMasterHelper = await new XmlRelationshipHelper().initialize(
      sourceArchive,
      `ppt/slideLayouts/_rels/slideLayout${slideLayoutId}.xml.rels`,
    );
    const layoutToMaster = layoutToMasterHelper.getTargetsByPrefix(
      '../slideMasters/slideMaster',
    );
    return layoutToMaster[0].number;
  }
}
