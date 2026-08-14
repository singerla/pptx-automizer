import { FileHelper } from '../helper/file-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import { ShapeTargetType, SourceIdentifier, Target } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IMaster } from '../interfaces/imaster';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import HasShapes from './has-shapes';
import { Layout } from './layout';
import { log } from '../helper/logger';

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

    this.sourcePath = PptPaths.slideMaster(this.sourceNumber);
    this.relsPath = PptPaths.slideMasterRels(this.sourceNumber);
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

    this.targetArchive = targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('masters');
    this.targetPath = PptPaths.slideMaster(this.targetNumber);
    this.targetRelsPath = PptPaths.slideMasterRels(this.targetNumber);
    this.sourceArchive = this.sourceTemplate.archive;

    log.info('Importing slideMaster ' + this.targetNumber);

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

    await this.flushTargetXml();
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
      PptPaths.slideMasterRels(this.targetNumber),
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
      PptPaths.theme(themeSourceId),
      this.targetArchive,
      PptPaths.theme(themeTargetId),
    );

    await this.contentTypes.appendThemeToContentType(themeTargetId);

    await XmlHelper.replaceAttribute(
      this.targetArchive,
      PptPaths.slideMasterRels(this.targetNumber),
      'Relationship',
      'Id',
      themeTarget.rId,
      PptPaths.relative.theme(themeTargetId),
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
      PptPaths.slideMaster(this.sourceNumber),
      this.targetArchive,
      PptPaths.slideMaster(this.targetNumber),
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.slideMasterRels(this.sourceNumber),
      this.targetArchive,
      PptPaths.slideMasterRels(this.targetNumber),
    );
  }
}
