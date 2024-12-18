import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { ShapeTargetType, SourceIdentifier, Target } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { IMaster } from '../interfaces/imaster';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import HasShapes from './has-shapes';
import { Layout } from './layout';
import { log } from '../helper/general-helper';

export class Master extends HasShapes implements IMaster {
  targetType: ShapeTargetType = 'slideMaster';
  key: string;

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
    sourceIdentifier: SourceIdentifier;
  }) {
    super(params);

    // ToDo analogue for slideMasters
    // this.sourceNumber = this.getSlideNumber(
    //   params.template,
    //   params.sourceIdentifier,
    // );

    this.sourceNumber = Number(params.sourceIdentifier);

    this.key = Master.getKey(this.sourceNumber, params.template.name);

    this.sourcePath = `ppt/slideMasters/slideMaster${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slideMasters/_rels/slideMaster${this.sourceNumber}.xml.rels`;
  }

  static getKey(slideLayoutNumber: number, templateName: string) {
    return slideLayoutNumber + '@' + templateName;
  }

  /**
   * Appends slide
   * @internal
   * @param targetTemplate
   * @returns append
   */
  async append(targetTemplate: RootPresTemplate): Promise<void> {
    this.targetTemplate = targetTemplate;

    this.targetArchive = await targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('masters');
    this.targetPath = `ppt/slideMasters/slideMaster${this.targetNumber}.xml`;
    this.targetRelsPath = `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`;
    this.sourceArchive = await this.sourceTemplate.archive;

    log('Importing slideMaster ' + this.targetNumber, 2);

    await this.copySlideMasterFiles();
    await this.copyRelatedLayouts();
    await this.copyRelatedContent();
    await this.addToPresentation();
    await this.copyThemeFiles();

    if (this.importElements.length) {
      await this.importedSelectedElements();
    }

    await this.applyModifications();
    await this.applyRelModifications();

    const info = this.targetTemplate.automizer.params.showIntegrityInfo;
    const assert = this.targetTemplate.automizer.params.showIntegrityInfo;
    await this.checkIntegrity(info, assert);

    await this.cleanSlide(this.targetPath);
  }

  async copyRelatedLayouts(): Promise<Target[]> {
    const targets = (await new XmlRelationshipHelper().initialize(
      this.targetArchive,
      `slideMaster${this.targetNumber}.xml.rels`,
      `ppt/slideMasters/_rels`,
      '../slideLayouts/slideLayout',
    )) as Target[];

    for (const target of targets) {
      const layout = new Layout({
        presentation: this.targetTemplate.automizer,
        template: this.sourceTemplate,
        sourceIdentifier: target.number,
        targetMaster: this.targetNumber,
      });

      await this.targetTemplate.appendLayout(layout);

      const layoutName = await layout.getName();

      this.targetTemplate.mapContents(
        'slideLayout',
        this.sourceTemplate.name,
        target.number,
        layout.targetNumber,
        layoutName,
      );

      target.updateTargetIndex(layout.targetNumber);
    }

    return targets;
  }

  async copyThemeFiles() {
    const targets = await XmlHelper.getRelationshipTargetsByPrefix(
      this.targetArchive,
      `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`,
      '../theme/theme',
    );

    if (!targets.length) {
      return;
    }

    const themeTarget = targets[0];

    const themeSourceId = themeTarget.number;
    const themeTargetId = this.targetTemplate.incrementCounter('themes');

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/theme/theme${themeSourceId}.xml`,
      this.targetArchive,
      `ppt/theme/theme${themeTargetId}.xml`,
    );

    await this.appendThemeToContentType(this.targetArchive, themeTargetId);

    await XmlHelper.replaceAttribute(
      this.targetArchive,
      `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`,
      'Relationship',
      'Id',
      themeTarget.rId,
      `../theme/theme${themeTargetId}.xml`,
      'Target',
    );
  }

  /**
   * Copy slide master files
   * @internal
   */
  async copySlideMasterFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/slideMasters/slideMaster${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/slideMasters/slideMaster${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/slideMasters/_rels/slideMaster${this.sourceNumber}.xml.rels`,
      this.targetArchive,
      `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`,
    );
  }

  appendThemeToContentType(
    rootArchive: IArchive,
    themeCount: string | number,
  ): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/theme/theme${themeCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.theme+xml`,
      }),
    );
  }
}
