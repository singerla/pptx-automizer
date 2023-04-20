import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import {
  AnalyzedElementType,
  FindElementSelector,
  FindElementStrategy,
  ImportedElement,
  ImportElement,
  ShapeModificationCallback,
  SlideModificationCallback,
  SourceIdentifier,
  StatusTracker,
} from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ElementType } from '../enums/element-type';
import {
  HelperElement,
  RelationshipAttribute,
  SlideListAttribute,
  XmlDocument,
} from '../types/xml-types';
import { Image } from '../shapes/image';
import { Chart } from '../shapes/chart';
import { GenericShape } from '../shapes/generic';
import { vd } from '../helper/general-helper';
import { ContentTracker } from '../helper/content-tracker';
import IArchive from '../interfaces/iarchive';

export class Slide implements ISlide {
  /**
   * Source template of slide
   * @internal
   */
  sourceTemplate: PresTemplate;
  /**
   * Target template of slide
   * @internal
   */
  targetTemplate: RootPresTemplate;
  /**
   * Target number of slide
   * @internal
   */
  targetNumber: number;
  /**
   * Source number of slide
   * @internal
   */
  sourceNumber: number;
  /**
   * Target archive of slide
   * @internal
   */
  targetArchive: IArchive;
  /**
   * Source archive of slide
   * @internal
   */
  sourceArchive: IArchive;
  /**
   * Source path of slide
   * @internal
   */
  sourcePath: string;
  /**
   * Target path of slide
   * @internal
   */
  targetPath: string;
  /**
   * Modifications  of slide
   * @internal
   */
  modifications: SlideModificationCallback[];
  /**
   * Import elements of slide
   * @internal
   */
  importElements: ImportElement[];
  /**
   * Rels path of slide
   * @internal
   */
  relsPath: string;
  /**
   * Root template of slide
   * @internal
   */
  rootTemplate: RootPresTemplate;
  /**
   * Root  of slide
   * @internal
   */
  root: IPresentationProps;
  /**
   * Target rels path of slide
   * @internal
   */
  targetRelsPath: string;
  status: StatusTracker;
  content: ContentTracker;
  /**
   * List of unsupported tags in slide xml
   * @internal
   */
  unsupportedTags = [
    'p:custDataLst',
    //'mc:AlternateContent',
    //'a14:imgProps',
  ];

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
    slideIdentifier: SourceIdentifier;
  }) {
    this.sourceTemplate = params.template;
    this.sourceNumber = this.getSlideNumber(
      params.template,
      params.slideIdentifier,
    );

    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`;

    this.modifications = [];
    this.importElements = [];

    this.status = params.presentation.status;
    this.content = params.presentation.content;
  }

  /**
   * Try to convert a given slide's creationId to corresponding slide number.
   * Used if automizer is run with useCreationIds: true
   * @internal
   * @param PresTemplate
   * @slideNumber SourceSlideIdentifier
   * @returns number
   */
  getSlideNumber(
    template: PresTemplate,
    slideIdentifier: SourceIdentifier,
  ): number {
    if (
      template.useCreationIds === true &&
      template.creationIds !== undefined
    ) {
      const matchCreationId = template.creationIds.find(
        (slideInfo) => slideInfo.id === Number(slideIdentifier),
      );

      if (matchCreationId) {
        return matchCreationId.number;
      }

      throw (
        'Could not find slide number for creationId: ' +
        slideIdentifier +
        '@' +
        template.name
      );
    }

    return slideIdentifier as number;
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
    await this.addSlideToPresentation();

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
    await this.cleanSlide();

    this.status.increment();
  }

  /**
   * Use another slide master
   * @param alias
   */
  useMaster(masterSelector: string, layoutSelector?: string): void {
    // TODO: get aliased master params and update current slide rels
  }

  /**
   * Modifies slide
   * @internal
   * @param callback
   */
  modify(callback: SlideModificationCallback): void {
    this.modifications.push(callback);
  }

  /**
   * Adds slide to presentation
   * @internal
   * @returns slide to presentation
   */
  async addSlideToPresentation(): Promise<void> {
    const relId = await XmlHelper.getNextRelId(
      this.targetArchive,
      'ppt/_rels/presentation.xml.rels',
    );
    await this.appendToSlideRel(this.targetArchive, relId, this.targetNumber);
    await this.appendToSlideList(this.targetArchive, relId);
    await this.appendSlideToContentType(this.targetArchive, this.targetNumber);
  }

  /**
   * Select and modify a single element on an added slide.
   * @param {string} selector - Element's name on the slide.
   * Should be a unique string defined on the "Selection"-pane within ppt.
   * @param {ShapeModificationCallback | ShapeModificationCallback[]} callback - One or more callback functions to apply.
   * Depending on the shape type (e.g. chart or table), different arguments will be passed to the callback.
   */
  modifyElement(
    selector: FindElementSelector,
    callback: ShapeModificationCallback | ShapeModificationCallback[],
  ): this {
    const presName = this.sourceTemplate.name;
    const slideNumber = this.sourceNumber;

    return this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'modify',
      callback,
    );
  }

  /**
   * Select, insert and (optionally) modify a single element to a slide.
   * @param {string} presName - Filename or alias name of the template presentation.
   * Must have been importet with Automizer.load().
   * @param {number} slideNumber - Slide number within the specified template to search for the required element.
   * @param {ShapeModificationCallback | ShapeModificationCallback[]} callback - One or more callback functions to apply.
   * Depending on the shape type (e.g. chart or table), different arguments will be passed to the callback.
   */
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): this {
    return this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'append',
      callback,
    );
  }

  /**
   * Remove a single element from slide.
   * @param {string} selector - Element's name on the slide.
   */
  removeElement(selector: FindElementSelector): this {
    const presName = this.sourceTemplate.name;
    const slideNumber = this.sourceNumber;

    return this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'remove',
      undefined,
    );
  }

  /**
   * Adds element to modifications list
   * @internal
   * @param presName
   * @param slideNumber
   * @param selector
   * @param mode
   * @param [callback]
   * @returns element to modifications list
   */
  private addElementToModificationsList(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    mode: string,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): this {
    this.importElements.push({
      presName,
      slideNumber,
      selector,
      mode,
      callback,
    });

    return this;
  }

  /**
   * Imported selected elements
   * @internal
   * @returns selected elements
   */
  async importedSelectedElements(): Promise<void> {
    for (const element of this.importElements) {
      const info = await this.getElementInfo(element);

      switch (info?.type) {
        case ElementType.Chart:
          await new Chart(info)[info.mode](
            this.targetTemplate,
            this.targetNumber,
          );
          break;
        case ElementType.Image:
          await new Image(info)[info.mode](
            this.targetTemplate,
            this.targetNumber,
          );
          break;
        case ElementType.Shape:
          await new GenericShape(info)[info.mode](
            this.targetTemplate,
            this.targetNumber,
          );
          break;
        default:
          break;
      }
    }
  }

  /**
   * Gets element info
   * @internal
   * @param importElement
   * @returns element info
   */
  async getElementInfo(importElement: ImportElement): Promise<ImportedElement> {
    const template = this.root.getTemplate(importElement.presName);

    const slideNumber =
      importElement.mode === 'append'
        ? this.getSlideNumber(template, importElement.slideNumber)
        : importElement.slideNumber;

    const sourcePath = `ppt/slides/slide${slideNumber}.xml`;

    const sourceArchive = await template.archive;
    const useCreationIds =
      template.useCreationIds === true && template.creationIds !== undefined;

    const { sourceElement, selector } = await this.findElementOnSlide(
      importElement.selector,
      sourceArchive,
      sourcePath,
      useCreationIds,
    );

    if (!sourceElement) {
      console.error(
        `Can't find element on slide ${slideNumber} in ${importElement.presName}: `,
      );
      console.log(importElement);
      return;
    }

    const appendElementParams = await this.analyzeElement(
      sourceElement,
      sourceArchive,
      slideNumber,
    );

    return {
      mode: importElement.mode,
      name: selector,
      hasCreationId: useCreationIds,
      sourceArchive,
      sourceSlideNumber: slideNumber,
      sourceElement,
      callback: importElement.callback,
      target: appendElementParams.target,
      type: appendElementParams.type,
    };
  }

  async findElementOnSlide(
    selector: FindElementSelector,
    sourceArchive: IArchive,
    sourcePath: string,
    useCreationIds: boolean,
  ): Promise<{
    sourceElement: XmlDocument;
    selector: string;
  }> {
    const strategies: FindElementStrategy[] = [];
    if (typeof selector === 'string') {
      if (useCreationIds) {
        strategies.push({
          mode: 'findByElementCreationId',
          selector: selector,
        });
      }
      strategies.push({
        mode: 'findByElementName',
        selector: selector,
      });
    } else if (selector.name) {
      strategies.push({
        mode: 'findByElementCreationId',
        selector: selector.creationId,
      });
      strategies.push({
        mode: 'findByElementName',
        selector: selector.name,
      });
    }

    for (const findElement of strategies) {
      const mode = findElement.mode;

      const sourceElement = await XmlHelper[mode](
        sourceArchive,
        sourcePath,
        findElement.selector,
      );

      if (sourceElement) {
        return { sourceElement, selector: findElement.selector };
      }
    }

    return { sourceElement: undefined, selector: JSON.stringify(selector) };
  }

  /**
   * Analyzes element
   * @internal
   * @param sourceElement
   * @param sourceArchive
   * @param slideNumber
   * @returns element
   */
  async analyzeElement(
    sourceElement: XmlDocument,
    sourceArchive: IArchive,
    slideNumber: number,
  ): Promise<AnalyzedElementType> {
    const isChart = sourceElement.getElementsByTagName('c:chart');
    if (isChart.length) {
      const target = await XmlHelper.getTargetByRelId(
        sourceArchive,
        slideNumber,
        sourceElement,
        'chart',
      );

      return {
        type: ElementType.Chart,
        target: target,
      } as AnalyzedElementType;
    }

    const isChartEx = sourceElement.getElementsByTagName('cx:chart');
    if (isChartEx.length) {
      const target = await XmlHelper.getTargetByRelId(
        sourceArchive,
        slideNumber,
        sourceElement,
        'chartEx',
      );

      return {
        type: ElementType.Chart,
        target: target,
      } as AnalyzedElementType;
    }

    const isImage = sourceElement.getElementsByTagName('p:nvPicPr');
    if (isImage.length) {
      return {
        type: ElementType.Image,
        target: await XmlHelper.getTargetByRelId(
          sourceArchive,
          slideNumber,
          sourceElement,
          'image',
        ),
      } as AnalyzedElementType;
    }

    return {
      type: ElementType.Shape,
    } as AnalyzedElementType;
  }

  /**
   * Applys modifications
   * @internal
   * @returns modifications
   */
  async applyModifications(): Promise<void> {
    for (const modification of this.modifications) {
      const xml = await XmlHelper.getXmlFromArchive(
        this.targetArchive,
        this.targetPath,
      );
      modification(xml);
      XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, xml);
    }
  }

  /**
   * Removes all unsupported tags from slide xml.
   * E.g. added relations & tags by Thinkcell cannot
   * be processed by pptx-automizer at the moment.
   * @internal
   */
  async cleanSlide(): Promise<void> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetPath,
    );

    this.unsupportedTags.forEach((tag) => {
      const drop = xml.getElementsByTagName(tag);
      const length = drop.length;
      if (length && length > 0) {
        XmlHelper.sliceCollection(drop, 0);
      }
    });
    XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, xml);
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
   * slideNote numbers differ from slide numbers if presentation
   * contains slides without notes. We need to find out
   * the proper enumeration of slideNote xml files.
   * @internal
   * @returns slide note file number
   */
  async getSlideNoteSourceNumber(): Promise<number | undefined> {
    const targets = await XmlHelper.getTargetsByRelationshipType(
      this.sourceArchive,
      `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    );

    if (targets.length) {
      const targetNumber = targets[0].file
        .replace('../notesSlides/notesSlide', '')
        .replace('.xml', '');
      return Number(targetNumber);
    }
  }

  /**
   * Copys slide note files
   * @internal
   * @returns slide note files
   */
  async copySlideNoteFiles(sourceNotesNumber: number): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/notesSlides/notesSlide${sourceNotesNumber}.xml`,
      this.targetArchive,
      `ppt/notesSlides/notesSlide${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/notesSlides/_rels/notesSlide${sourceNotesNumber}.xml.rels`,
      this.targetArchive,
      `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`,
    );
  }

  /**
   * Updates slide note file
   * @internal
   * @returns slide note file
   */
  async updateSlideNoteFile(sourceNotesNumber: number): Promise<void> {
    await XmlHelper.replaceAttribute(
      this.targetArchive,
      `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`,
      'Relationship',
      'Target',
      `../slides/slide${this.sourceNumber}.xml`,
      `../slides/slide${this.targetNumber}.xml`,
    );

    await XmlHelper.replaceAttribute(
      this.targetArchive,
      `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`,
      'Relationship',
      'Target',
      `../notesSlides/notesSlide${sourceNotesNumber}.xml`,
      `../notesSlides/notesSlide${this.targetNumber}.xml`,
    );
  }

  /**
   * Appends to slide rel
   * @internal
   * @param rootArchive
   * @param relId
   * @param slideCount
   * @returns to slide rel
   */
  appendToSlideRel(
    rootArchive: IArchive,
    relId: string,
    slideCount: number,
  ): Promise<HelperElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/_rels/presentation.xml.rels`,
      parent: (xml: XmlDocument) =>
        xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`,
        Target: `slides/slide${slideCount}.xml`,
      } as RelationshipAttribute,
    });
  }

  /**
   * Appends a new slide to slide list in presentation.xml.
   * If rootArchive has no slides, a new node will be created.
   * "id"-attribute of 'p:sldId'-element must be greater than 255.
   * @internal
   * @param rootArchive
   * @param relId
   * @returns to slide list
   */
  appendToSlideList(
    rootArchive: IArchive,
    relId: string,
  ): Promise<HelperElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      assert: async (xml: XmlDocument) => {
        if (xml.getElementsByTagName('p:sldIdLst').length === 0) {
          XmlHelper.insertAfter(
            xml.createElement('p:sldIdLst'),
            xml.getElementsByTagName('p:sldMasterIdLst')[0],
          );
        }
      },
      parent: (xml: XmlDocument) => xml.getElementsByTagName('p:sldIdLst')[0],
      tag: 'p:sldId',
      attributes: {
        'r:id': relId,
      } as SlideListAttribute,
    });
  }

  /**
   * Appends slide to content type
   * @internal
   * @param rootArchive
   * @param slideCount
   * @returns slide to content type
   */
  appendSlideToContentType(
    rootArchive: IArchive,
    slideCount: number,
  ): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`,
      }),
    );
  }

  /**
   * Appends notes to content type
   * @internal
   * @param rootArchive
   * @param slideCount
   * @returns notes to content type
   */
  appendNotesToContentType(
    rootArchive: IArchive,
    slideCount: number,
  ): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/notesSlides/notesSlide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml`,
      }),
    );
  }

  /**
   * Copys related content
   * @internal
   * @returns related content
   */
  async copyRelatedContent(): Promise<void> {
    const charts = await Chart.getAllOnSlide(this.sourceArchive, this.relsPath);

    for (const chart of charts) {
      await new Chart({
        mode: 'append',
        target: chart,
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }

    const images = await Image.getAllOnSlide(this.sourceArchive, this.relsPath);
    for (const image of images) {
      await new Image({
        mode: 'append',
        target: image,
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }
  }
}
