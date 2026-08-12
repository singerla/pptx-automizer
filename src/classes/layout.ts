import { FileHelper } from '../helper/file-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import { ShapeTargetType, SourceIdentifier, Target } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import HasShapes from './has-shapes';
import { ILayout } from '../interfaces/ilayout';
import { log } from '../helper/logger';

export class Layout extends HasShapes implements ILayout {
  targetType: ShapeTargetType = 'slideLayout';
  targetMaster: number;

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
    sourceIdentifier: SourceIdentifier;
    targetMaster: number;
  }) {
    super(params);

    this.sourceNumber = Number(params.sourceIdentifier);
    this.targetMaster = params.targetMaster;

    this.sourcePath = PptPaths.slideLayout(this.sourceNumber);
    this.relsPath = PptPaths.slideLayoutRels(this.sourceNumber);
  }

  /**
   * Appends slideLayout
   * @internal
   * @param targetTemplate
   * @returns append
   */
  async append(targetTemplate: RootPresTemplate): Promise<void> {
    this.targetTemplate = targetTemplate;

    this.targetArchive = targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('layouts');
    this.targetPath = PptPaths.slideLayout(this.targetNumber);
    this.targetRelsPath = PptPaths.slideLayoutRels(this.targetNumber);
    this.sourceArchive = this.sourceTemplate.archive;

    log.info('Importing slideLayout ' + this.targetNumber);

    await this.copySlideLayoutFiles();
    await this.copyRelatedContent();
    await this.addToPresentation();
    await this.updateRelation();

    await this.cleanSlide(this.targetPath);
    await this.cleanRelations(this.targetRelsPath);
    await this.checkIntegrity(true, true);
  }

  /**
   * Copys slide layout files
   * @internal
   */
  async copySlideLayoutFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.slideLayout(this.sourceNumber),
      this.targetArchive,
      PptPaths.slideLayout(this.targetNumber),
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.slideLayoutRels(this.sourceNumber),
      this.targetArchive,
      PptPaths.slideLayoutRels(this.targetNumber),
    );
  }

  async updateRelation() {
    const layoutToMaster = (await new XmlRelationshipHelper().initialize(
      this.targetArchive,
      `slideLayout${this.targetNumber}.xml.rels`,
      `ppt/slideLayouts/_rels`,
      '../slideMasters/slideMaster',
    )) as Target[];

    layoutToMaster[0].updateTargetIndex(this.targetMaster);
  }

  async getName(): Promise<string> {
    const slideLayoutXml = await XmlHelper.getXmlFromArchive(
      this.sourceArchive,
      PptPaths.slideLayout(this.sourceNumber),
    );

    const layout = slideLayoutXml.getElementsByTagName('p:cSld')?.item(0);
    if (layout) {
      const name = layout.getAttribute('name');
      return name;
    }
  }
}
