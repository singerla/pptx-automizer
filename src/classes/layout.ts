import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { ShapeTargetType, SourceIdentifier, Target } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import HasShapes from './has-shapes';
import { ILayout } from '../interfaces/ilayout';
import { log } from '../helper/general-helper';

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

    this.sourcePath = `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slideLayouts/_rels/slideLayout${this.sourceNumber}.xml.rels`;
  }

  /**
   * Appends slideLayout
   * @internal
   * @param targetTemplate
   * @returns append
   */
  async append(targetTemplate: RootPresTemplate): Promise<void> {
    this.targetTemplate = targetTemplate;

    this.targetArchive = await targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('layouts');
    this.targetPath = `ppt/slideLayouts/slideLayout${this.targetNumber}.xml`;
    this.targetRelsPath = `ppt/slideLayouts/_rels/slideLayout${this.targetNumber}.xml.rels`;
    this.sourceArchive = await this.sourceTemplate.archive;

    log('Importing slideLayout ' + this.targetNumber, 2);

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
      `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/slideLayouts/slideLayout${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/slideLayouts/_rels/slideLayout${this.sourceNumber}.xml.rels`,
      this.targetArchive,
      `ppt/slideLayouts/_rels/slideLayout${this.targetNumber}.xml.rels`,
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
      `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`,
    );

    const layout = slideLayoutXml.getElementsByTagName('p:cSld')?.item(0);
    if (layout) {
      const name = layout.getAttribute('name');
      return name;
    }
  }
}
