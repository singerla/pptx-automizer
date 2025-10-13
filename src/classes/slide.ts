import { FileHelper } from '../helper/file-helper';
import {
  ShapeModificationCallback,
  ShapeTargetType,
  SourceIdentifier,
} from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { GeneralHelper, last, vd } from '../helper/general-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';
import { Master } from './master';
import ModifyPresentationHelper from '../helper/modify-presentation-helper';
import {
  ElementInfo,
  PlaceholderInfo,
  PlaceholderMappingResult,
  XmlElement,
} from '../types/xml-types';
import XmlPlaceholderHelper from '../helper/xml-placeholder-helper';

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

    const placeholderTypes = await this.parsePlaceholders();

    if (this.importElements.length) {
      await this.importedSelectedElements();
    }

    await this.applyModifications();
    await this.applyRelModifications();

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
   * Merges slide content into a specified slide layout by mapping placeholders.
   * This method automatically handles placeholder matching and repositioning of elements
   * that don't have corresponding placeholders in the target layout.
   *
   * @param targetLayout - Name or identifier of the target slide layout to merge into
   * @param targetPlaceholders - Array of placeholder information from the target layout
   * @returns Promise<this> - Returns the slide instance for method chaining
   */
  async mergeIntoSlideLayout(
    targetLayout: string,
    targetPlaceholders: PlaceholderInfo[],
  ): Promise<this> {
    // Step 1: Apply the target layout to this slide
    this.useSlideLayout(targetLayout);

    // Step 2: Gather source layout information and slide elements
    const sourceLayoutInfo = await this.getSourceLayoutInfo();
    const slideElements = await this.getAllElements([]);

    // Step 3: Initialize tracking arrays for placeholder mapping
    const placeholderMappingResult = this.initializePlaceholderMapping();

    // Step 4: Perform initial placeholder matching
    this.performInitialPlaceholderMatching(
      slideElements,
      targetPlaceholders,
      placeholderMappingResult,
    );

    // Step 5: Handle unmatched elements with alternative matching
    this.handleUnmatchedElements(
      placeholderMappingResult.unmatchedElements,
      targetPlaceholders,
      placeholderMappingResult.usedPlaceholders,
    );

    // Step 6: Clean up remaining unmatched placeholders
    this.cleanupUnmatchedPlaceholders(
      placeholderMappingResult.unmatchedElements,
      sourceLayoutInfo.placeholders,
    );

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
   * Initializes the data structures needed for placeholder mapping.
   *
   * @returns PlaceholderMappingResult - Object containing arrays for tracking mappings
   * @private
   */
  private initializePlaceholderMapping(): PlaceholderMappingResult {
    return {
      usedPlaceholders: [],
      unmatchedElements: [],
    };
  }

  /**
   * Performs the initial placeholder matching between source elements and target placeholders.
   * Elements with exact placeholder type matches are processed first.
   *
   * @param elements - Array of slide elements to process
   * @param targetPlaceholders - Array of available placeholders in target layout
   * @param mappingResult - Object to track mapping results
   * @private
   */
  private performInitialPlaceholderMatching(
    elements: ElementInfo[],
    targetPlaceholders: PlaceholderInfo[],
    mappingResult: PlaceholderMappingResult,
  ): void {
    this.pushPlaceholderUsage(
      elements,
      mappingResult.unmatchedElements,
      targetPlaceholders,
      mappingResult.usedPlaceholders,
    );
  }

  /**
   * Handles elements that couldn't be matched in the initial pass by finding
   * alternative placeholder matches using best-fit algorithms.
   *
   * @param unmatchedElements - Array of elements that need alternative matches
   * @param targetPlaceholders - Array of available placeholders in target layout
   * @param usedPlaceholders - Array of placeholders already assigned
   * @private
   */
  private handleUnmatchedElements(
    unmatchedElements: ElementInfo[],
    targetPlaceholders: PlaceholderInfo[],
    usedPlaceholders: PlaceholderInfo[],
  ): void {
    // Create a copy to avoid modifying array during iteration
    const elementsToProcess = [...unmatchedElements];

    elementsToProcess.forEach((element) => {
      const bestAlternativeMatch = XmlPlaceholderHelper.findBestTargetPlaceholderAlternative(
        element,
        targetPlaceholders,
        usedPlaceholders,
      );

      if (bestAlternativeMatch) {
        this.applyAlternativePlaceholderMatch(
          element,
          bestAlternativeMatch,
          usedPlaceholders,
        );
        this.removeElementFromUnmatched(element, unmatchedElements);
      }
    });
  }

  /**
   * Applies an alternative placeholder match to an element.
   *
   * @param element - Element to apply the match to
   * @param placeholder - Placeholder to match with
   * @param usedPlaceholders - Array to track used placeholders
   * @private
   */
  private applyAlternativePlaceholderMatch(
    element: ElementInfo,
    placeholder: PlaceholderInfo,
    usedPlaceholders: PlaceholderInfo[],
  ): void {
    this.applyPlaceholder(element, placeholder, usedPlaceholders);
  }

  /**
   * Removes an element from the unmatched elements array.
   *
   * @param element - Element to remove
   * @param unmatchedElements - Array to remove from
   * @private
   */
  private removeElementFromUnmatched(
    element: ElementInfo,
    unmatchedElements: ElementInfo[],
  ): void {
    const index = unmatchedElements.indexOf(element);
    if (index > -1) {
      unmatchedElements.splice(index, 1);
    }
  }

  /**
   * Cleans up elements that still don't have placeholder matches by removing
   * their placeholder properties and applying fallback positioning.
   *
   * @param unmatchedElements - Elements that couldn't be matched
   * @param sourcePlaceholders - Original placeholders from source layout for fallback
   * @private
   */
  private cleanupUnmatchedPlaceholders(
    unmatchedElements: ElementInfo[],
    sourcePlaceholders: PlaceholderInfo[],
  ): void {
    unmatchedElements.forEach((element) => {
      this.clearUnmatchedPlaceholder(element, sourcePlaceholders);
    });
  }

  pushPlaceholderUsage(
    elements: ElementInfo[],
    unmatchedPhElements: ElementInfo[],
    targetPlaceholders: PlaceholderInfo[],
    usedPlaceholders: PlaceholderInfo[],
  ): void {
    elements.forEach((element: ElementInfo) => {
      if (element.placeholder?.type) {
        const matchesPlaceholder = this.applyPlaceholderToElement(
          targetPlaceholders,
          element.placeholder.type,
          usedPlaceholders,
          element,
        );
        if (!matchesPlaceholder) {
          unmatchedPhElements.push(element);
        }
      }
    });
  }

  clearUnmatchedPlaceholder(
    element: ElementInfo,
    sourcePlaceholders: PlaceholderInfo[],
  ) {
    const fallbackPh = sourcePlaceholders.find(
      (ph) => ph.idx === element.placeholder.idx,
    );
    const fallbackPosition = fallbackPh?.position || {
      x: 1000,
      y: 1000,
      cx: 5000000,
      cy: 1000000,
    };

    const callback = (element) => {
      XmlPlaceholderHelper.removePlaceholder(element, fallbackPosition);
    };
    this.postApplyModification(element, callback);
  }

  applyPlaceholderToElement(
    layoutPlaceholders: PlaceholderInfo[],
    forceType: string,
    usedPlaceholders: PlaceholderInfo[],
    element: ElementInfo,
  ): PlaceholderInfo {
    const unusedPlaceholders = layoutPlaceholders.filter(
      (ph) => !usedPlaceholders.includes(ph),
    );

    const matchPlaceholders = unusedPlaceholders.filter((ph) => {
      return ph.type === forceType;
    });

    if (matchPlaceholders.length) {
      const bestMatch = XmlPlaceholderHelper.findBestMatchingPlaceholder(
        element,
        matchPlaceholders,
      );
      this.applyPlaceholder(element, bestMatch, usedPlaceholders);
      return bestMatch;
    }
    return null;
  }

  applyPlaceholder(
    element: ElementInfo,
    bestMatch: PlaceholderInfo,
    usedPlaceholders: PlaceholderInfo[],
  ) {
    const callback = (element: XmlElement) => {
      XmlPlaceholderHelper.setPlaceholderDefaults(element, bestMatch);
    };
    this.postApplyModification(element, callback);
    usedPlaceholders.push(bestMatch);
  }

  postApplyModification(
    element: ElementInfo,
    callback: ShapeModificationCallback,
  ) {
    const selector = {
      creationId: element.creationId,
      nameIdx: element.nameIdx,
      name: element.name,
    };

    const alreadyModified = this.getAlreadyModifiedElement(selector);
    if (alreadyModified) {
      alreadyModified.callback = GeneralHelper.arrayify(
        alreadyModified.callback,
      );
      alreadyModified.callback.push(callback);
    } else {
      this.modifyElement(selector, callback);
    }
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
