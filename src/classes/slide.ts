import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { last, vd } from '../helper/general-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';

export class Slide extends HasShapes implements ISlide {
  targetType: ShapeTargetType = 'slide';

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
    slideIdentifier: SourceIdentifier;
  }) {
    super(params);

    this.sourceNumber = this.getSlideNumber(
      params.template,
      params.slideIdentifier,
    );

    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`;
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
    this.targetNumber = targetTemplate.incrementCounter('slides');
    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`;
    this.targetRelsPath = `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`;
    this.sourceArchive = await this.sourceTemplate.archive;

    this.status.info = 'Appending slide ' + this.targetNumber;

    await this.copySlideFiles();
    await this.copyRelatedContent();
    await this.addToPresentation();

    const sourceNotesNumber = await this.getSlideNoteSourceNumber();
    if (sourceNotesNumber) {
      await this.copySlideNoteFiles(sourceNotesNumber);
      await this.updateSlideNoteFile(sourceNotesNumber);
      await this.appendNotesToContentType(
        this.targetArchive,
        this.targetNumber,
      );
    }

    if (this.importElements.length) {
      await this.importedSelectedElements();
    }

    await this.applyModifications();
    await this.applyRelModifications();

    const info = this.targetTemplate.automizer.params.showIntegrityInfo;
    const assert = this.targetTemplate.automizer.params.showIntegrityInfo;
    await this.checkIntegrity(info, assert);

    await this.cleanSlide(this.targetPath);

    this.status.increment();
  }

  /**
   * Use another slide layout.
   * @param targetLayoutId
   */
  useSlideLayout(layoutId?: number | string): this {
    this.relModifications.push(async (slideRelXml) => {
      let targetLayoutId;

      if (typeof layoutId === 'string') {
        targetLayoutId = await this.useNamedSlideLayout(layoutId as string);

        if (!targetLayoutId) {
          layoutId = null;
        }
      }

      if (!layoutId || typeof layoutId === 'number') {
        targetLayoutId = await this.useIndexedSlideLayout(layoutId as number);
      }

      const slideLayouts = new XmlRelationshipHelper(slideRelXml)
        .readTargets()
        .getTargetsByPrefix('../slideLayouts/slideLayout');

      if (slideLayouts.length) {
        slideLayouts[0].updateTargetIndex(targetLayoutId as number);
      }
    });

    return this;
  }

  /**
   * Find another slide layout by name.
   * @param targetLayoutId
   */
  async useNamedSlideLayout(targetLayoutName: string): Promise<number> {
    const templateName = this.sourceTemplate.name;
    const sourceLayoutId = await XmlRelationshipHelper.getSlideLayoutNumber(
      this.sourceArchive,
      this.sourceNumber,
    );

    await this.autoImportSourceSlideMaster(templateName, sourceLayoutId);

    const alreadyImported = this.targetTemplate.getNamedMappedContent(
      'slideLayout',
      targetLayoutName,
    );

    if (!alreadyImported) {
      console.error(
        'Could not find "' +
          targetLayoutName +
          '"@' +
          templateName +
          '@' +
          'sourceLayoutId:' +
          sourceLayoutId,
      );
    }

    return alreadyImported?.targetId;
  }

  /**
   * Use another slide layout by index or detect original index.
   * @param targetLayoutId
   */
  async useIndexedSlideLayout(targetLayoutIndex?: number): Promise<number> {
    if (!targetLayoutIndex) {
      const sourceLayoutId = await XmlRelationshipHelper.getSlideLayoutNumber(
        this.sourceArchive,
        this.sourceNumber,
      );

      const templateName = this.sourceTemplate.name;
      const alreadyImported = this.targetTemplate.getMappedContent(
        'slideLayout',
        templateName,
        sourceLayoutId,
      );

      if (alreadyImported) {
        return alreadyImported.targetId;
      } else {
        return await this.autoImportSourceSlideMaster(
          templateName,
          sourceLayoutId,
        );
      }
    }
    return targetLayoutIndex;
  }

  async autoImportSourceSlideMaster(
    templateName: string,
    sourceLayoutId: number,
  ) {
    const sourceMasterId = await XmlRelationshipHelper.getSlideMasterNumber(
      this.sourceArchive,
      sourceLayoutId,
    );
    await this.targetTemplate.automizer.addMaster(templateName, sourceMasterId);

    const previouslyAddedMaster = last<IMaster>(this.targetTemplate.masters);

    await this.targetTemplate
      .appendMasterSlide(previouslyAddedMaster)
      .catch((e) => {
        throw e;
      });

    const alreadyImported = this.targetTemplate.getMappedContent(
      'slideLayout',
      templateName,
      sourceLayoutId,
    );

    return alreadyImported.targetId;
  }

  /**
   * Apply modifications to slide relations
   * @internal
   * @returns modifications
   */
  async applyRelModifications(): Promise<void> {
    await XmlHelper.modifyXmlInArchive(
      this.targetArchive,
      `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`,
      this.relModifications,
    );
  }

  /**
   * Copys slide files
   * @internal
   */
  async copySlideFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/slides/slide${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/slides/slide${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`,
      this.targetArchive,
      `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`,
    );
  }
}
