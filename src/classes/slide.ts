import { FileHelper } from '../helper/file-helper';
import { XmlDocument } from '../types/xml-types';
import { PptPaths } from '../helper/ppt-paths';
import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { last } from '../helper/general-helper';
import { log } from '../helper/logger';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';
import { Master } from './master';
import ModifyPresentationHelper from '../helper/modify-presentation-helper';
import XmlPlaceholderHelper from '../helper/xml-placeholder-helper';

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

    this.sourcePath = PptPaths.slide(this.sourceNumber);
    this.relsPath = PptPaths.slideRels(this.sourceNumber);
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
    this.targetNumber = targetTemplate.incrementCounter('slides');
    this.targetPath = PptPaths.slide(this.targetNumber);
    this.targetRelsPath = PptPaths.slideRels(this.targetNumber);
    this.sourceArchive = this.sourceTemplate.archive;

    this.status.info = 'Appending slide ' + this.targetNumber;

    await this.copySlideFiles();
    await this.copyRelatedContent();
    await this.addToPresentation();
    await this.notes.copySlideNotes();

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

    await this.flushTargetXml();

    this.status.increment();
  }

  /**
   * Additionally flushes the slide's copied notesSlide part, which shares
   * the slide's target number.
   * @internal
   */
  async flushTargetXml(): Promise<void> {
    await super.flushTargetXml();
    await this.targetArchive.flushXml(PptPaths.notesSlide(this.targetNumber));
    await this.targetArchive.flushXml(
      PptPaths.notesSlideRels(this.targetNumber),
    );
  }

  /**
   * Use another slide layout.
   * @param layoutId
   */
  useSlideLayout(layoutId?: number | string): this {
    this.relModifications.push(async (slideRelXml: XmlDocument) => {
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
        log.warn('Unable to use slide layout ' + layoutId);
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
    // Disabling concurring cleanup function for this slide:
    this.cleanupPlaceholders = false

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
      log.error(
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
      PptPaths.slide(this.sourceNumber),
      this.targetArchive,
      PptPaths.slide(this.targetNumber),
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.slideRels(this.sourceNumber),
      this.targetArchive,
      PptPaths.slideRels(this.targetNumber),
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
