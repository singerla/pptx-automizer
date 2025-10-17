import { FileHelper } from '../helper/file-helper';
import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { last, Logger, vd } from '../helper/general-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';
import { Master } from './master';
import ModifyPresentationHelper from '../helper/modify-presentation-helper';
import XmlPlaceholderHelper from '../helper/xml-placeholder-helper';
import { XmlSlideHelper } from '../helper/xml-slide-helper';
import { XmlTemplateHelper } from '../helper/xml-template-helper';

export class Slide extends HasShapes implements ISlide {
  targetType: ShapeTargetType = 'slide';
  private targetLayoutId: number;

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

    const placeholderTypes = await this.parsePlaceholders();

    await this.applyRelModifications();
    await this.applyPreparations();

    if (this.importElements.length) {
      await this.importedSelectedElements();
    }

    await this.applyModifications();

    const info = this.targetTemplate.automizer.params.showIntegrityInfo;
    const assert = this.targetTemplate.automizer.params.showIntegrityInfo;
    await this.checkIntegrity(info, assert);

    await this.cleanSlide(this.targetPath, placeholderTypes);

    this.status.increment();
  }

  /**
   * Use another slide layout.
   * @param layoutId
   */
  useSlideLayout(layoutId?: number | string): this {
    this.relModifications.push(async (slideRelXml: XMLDocument) => {
      let targetLayoutId: number;

      if (typeof layoutId === 'string') {
        targetLayoutId = await this.useNamedSlideLayout(layoutId as string);

        if (!targetLayoutId) {
          layoutId = null;
        }
      }

      if (!layoutId || typeof layoutId === 'number') {
        targetLayoutId = await this.useIndexedSlideLayout(layoutId as number);
      }

      if (targetLayoutId) {
        this.targetLayoutId = targetLayoutId
        const slideLayouts = new XmlRelationshipHelper(slideRelXml)
          .readTargets()
          .getTargetsByPrefix('../slideLayouts/slideLayout');

        if (slideLayouts.length) {
          slideLayouts[0].updateTargetIndex(targetLayoutId as number);
        }
      } else {
        Logger.log('Unable to use slide layout ' + layoutId, 0);
      }
    });

    return this;
  }

  /**
   * Merges slide content into a specified slide layout by mapping placeholders.
   * This method automatically handles placeholder matching and repositioning of elements
   * that don't have corresponding placeholders in the target layout.
   *
   * @param targetFileName
   * @param targetLayout - Name or identifier of the target slide layout to merge into
   * @returns Promise<this> - Returns the slide instance for method chaining
   */
  mergeIntoSlideLayout(targetLayout: number | string): this {
    this.useSlideLayout(targetLayout)

    this.prepare(async (_) => {
      const slideHelper = await this.getSlideHelperInstance(
        this.targetArchive,
        this.targetPath,
        this.targetNumber
      )
      const slideLayout = await slideHelper.getSlideLayout()
      const targetPlaceholders = slideLayout.placeholders || [];
      const sourceLayoutInfo = await this.getSourceLayoutInfo();
      const slideElements = await this.getAllElements([], targetPlaceholders);

      new XmlPlaceholderHelper(
        this,
        slideElements,
        sourceLayoutInfo,
        targetPlaceholders,
      ).run();
    })

    return this;
  }

  /**
   * Retrieves information about the source slide layout.
   *
   * @returns Promise<{placeholders: PlaceholderInfo[]}> Source layout information
   * @private
   */
  private async getSourceLayoutInfo() {
    const slideHelper = await this.getSlideHelper();
    const sourceLayout = await slideHelper.getSlideLayout();
    return sourceLayout;
  }

  /**
   * Find another slide layout by name.
   * @param targetLayoutName
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
   * @param targetLayoutIndex
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
    const key = Master.getKey(sourceMasterId, templateName);

    if (!this.targetTemplate.masters.find((master) => master.key === key)) {
      await this.targetTemplate.automizer.addMaster(
        templateName,
        sourceMasterId,
      );

      const previouslyAddedMaster = last<IMaster>(this.targetTemplate.masters);

      await this.targetTemplate
        .appendMasterSlide(previouslyAddedMaster)
        .catch((e) => {
          throw e;
        });
    }

    const alreadyImported = this.targetTemplate.getMappedContent(
      'slideLayout',
      templateName,
      sourceLayoutId,
    );

    return alreadyImported?.targetId;
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

  /**
   * Remove a slide from presentation's slide list.
   * ToDo: Find the current count for this slide;
   * ToDo: this.targetNumber is undefined at this point.
   */
  remove(slide: number): void {
    this.root.modify(ModifyPresentationHelper.removeSlides([slide]));
  }
}
