import { FileHelper } from '../helper/file-helper';
import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { last, vd } from '../helper/general-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';
import { Master } from './master';
import ModifyPresentationHelper from '../helper/modify-presentation-helper';
import { ElementInfo, PlaceholderInfo, SlideInfo } from '../types/xml-types';
import XmlPlaceholderHelper from '../helper/xml-placeholder-helper';
import { XmlHelper } from '../helper/xml-helper';

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

  async mergeIntoSlideLayout(
    targetLayout: string,
    slidesInfo: SlideInfo[],
  ): Promise<this> {
    this.useSlideLayout(targetLayout);

    const elements = await this.getAllElements();

    const layoutPlaceholders =
      slidesInfo.find((slide) => slide.info.layoutName === targetLayout)?.info
        .layoutPlaceholders || [];

    // vd(layoutPlaceholders)

    const usedPlaceholders: number[] = [];
    const unmatchedPhElements: ElementInfo[] = [];
    elements.forEach((element: ElementInfo) => {
      if (element.placeholder) {
        if (element.placeholder.type) {
          const matchesPlaceholder = this.applyPlaceholderToElement(
            layoutPlaceholders,
            element.placeholder.type,
            usedPlaceholders,
            element,
          );

          if (!matchesPlaceholder) {
            unmatchedPhElements.push(element);
          }
        } else {
          unmatchedPhElements.push(element);
        }
      }
    });

    unmatchedPhElements.forEach((element) => {
      const forceType = !element.placeholder.type
        ? 'body'
        : element.placeholder.type;

      const matchesPlaceholder = this.applyPlaceholderToElement(
        layoutPlaceholders,
        forceType,
        usedPlaceholders,
        element,
      );

      if (!matchesPlaceholder) {
        const forceType = element.placeholder.type === 'title'
          ? 'ctrTitle'
          : 'subTitle';

        const matchesPlaceholder2 = this.applyPlaceholderToElement(
          layoutPlaceholders,
          forceType,
          usedPlaceholders,
          element,
        );

        if (!matchesPlaceholder2) {
          this.modifyElement(
            {
              creationId: element.creationId,
              name: element.name,
            },
            (element) => {
              XmlPlaceholderHelper.removePlaceholder(element)
            },
          );
        }
      }
    });

    return this;
  }

  applyPlaceholderToElement(
    layoutPlaceholders: PlaceholderInfo[],
    forceType: string,
    usedPlaceholders: number[],
    element: ElementInfo,
  ): boolean {
    const unusedPlaceholders = layoutPlaceholders.filter(
      (ph) => !usedPlaceholders.includes(ph.idx),
    );

    const matchPlaceholders = unusedPlaceholders.filter((ph) => {
      return ph.type === forceType;
    });

    if (matchPlaceholders.length) {
      const matchPlaceholder = XmlPlaceholderHelper.findBestTargetPlaceholder(
        element,
        matchPlaceholders,
      );
      usedPlaceholders.push(matchPlaceholder.idx);

      this.modifyElement(
        {
          creationId: element.creationId,
          name: element.name,
        },
        (element) => {
          XmlPlaceholderHelper.resetPlaceholderToDefaults(
            element,
            matchPlaceholder,
          );
        },
      );
      return true;
    }
    return false;
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
