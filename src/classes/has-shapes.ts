import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import IArchive from '../interfaces/iarchive';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import {
  AnalyzedElementType,
  AutomizerParams,
  FindElementSelector,
  FindElementStrategy,
  ImportedElement,
  ImportElement,
  ShapeModificationCallback,
  ShapeTargetType,
  SlideModificationCallback,
  SourceIdentifier,
  StatusTracker,
} from '../types/types';
import { ContentTracker } from '../helper/content-tracker';
import { vd } from '../helper/general-helper';
import {
  HelperElement,
  RelationshipAttribute,
  SlideListAttribute,
  XmlDocument,
} from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { FileHelper } from '../helper/file-helper';
import { Chart } from '../shapes/chart';
import { Image } from '../shapes/image';
import { ElementType } from '../enums/element-type';
import { GenericShape } from '../shapes/generic';
import { XmlSlideHelper } from '../helper/xml-slide-helper';

export default class HasShapes {
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
   * Root template of slide
   * @internal
   */
  modifications: SlideModificationCallback[];
  /**
   * Modifications of slide relations
   * @internal
   */
  relModifications: SlideModificationCallback[];
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
   * Target rels path of slide
   * @internal
   */
  targetRelsPath: string;
  /**
   * Root  of slide
   * @internal
   */
  root: IPresentationProps;
  status: StatusTracker;
  content: ContentTracker;
  /**
   * List of unsupported tags in slide xml
   * @internal
   */
  unsupportedTags = [
    'p:custDataLst',
    // 'mc:AlternateContent',
    //'a14:imgProps',
  ];
  targetType: ShapeTargetType;
  params: AutomizerParams;

  constructor(params) {
    this.sourceTemplate = params.template;

    this.modifications = [];
    this.relModifications = [];
    this.importElements = [];

    this.status = params.presentation.status;
    this.content = params.presentation.content;
  }

  /**
 * Asynchronously retrieves all text element IDs from the slide.
 * @returns {Promise<string[]>} A promise that resolves to an array of text element IDs.
 */
async getAllTextElementIds(): Promise<string[]> {
  try {
    const template = this.sourceTemplate
    // Retrieve the slide XML data
    const slideXml = await XmlHelper.getXmlFromArchive(
      template.archive,
      this.sourcePath,
    );
    // Initialize the XmlSlideHelper
    const xmlSlideHelper = new XmlSlideHelper(slideXml);

    // Get all text element IDs 
    const textElementIds = xmlSlideHelper.getAllTextElementIds(template.useCreationIds || false);

    return textElementIds;
  } catch (error) {
    // Log the error message and return an empty array, none of the others seem to have any error handling.. so not sure whats best throw actual Error.., console.error, something else?
   /*  console.error(error.message);
    return []; */
    throw new Error(error.message)
  }
}


  /**
   * Push modifications list
   * @internal
   * @param callback
   */
  modify(callback: SlideModificationCallback): void {
    this.modifications.push(callback);
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
   * ToDo: Implement creationIds as well for slideMasters
   *
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
   * Imported selected elements
   * @internal
   * @returns selected elements
   */
  async importedSelectedElements(): Promise<void> {
    for (const element of this.importElements) {
      const info = await this.getElementInfo(element);

      switch (info?.type) {
        case ElementType.Chart:
          await new Chart(info, this.targetType)[info.mode](
            this.targetTemplate,
            this.targetNumber,
            this.targetType,
          );
          break;
        case ElementType.Image:
          await new Image(info, this.targetType)[info.mode](
            this.targetTemplate,
            this.targetNumber,
            this.targetType,
          );
          break;
        case ElementType.Shape:
          await new GenericShape(info, this.targetType)[info.mode](
            this.targetTemplate,
            this.targetNumber,
            this.targetType,
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

    let sourcePath = `ppt/slides/slide${slideNumber}.xml`;

    if (this.targetType === 'slideMaster') {
      // It is possible to import shapes from loaded presentations,
      // as well as to modify an existing shape on current slideMaster
      sourcePath =
        importElement.mode === 'append'
          ? `ppt/slides/slide${slideNumber}.xml`
          : `ppt/slideMasters/slideMaster${slideNumber}.xml`;
    }

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

  /**
   * @param selector
   * @param sourceArchive
   * @param sourcePath
   * @param useCreationIds
   */
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

  async checkIntegrity(info: boolean, assert: boolean): Promise<void> {
    if (info || assert) {
      const masterRels = (await new XmlRelationshipHelper().initialize(
        this.targetArchive,
        `${this.targetType}${this.targetNumber}.xml.rels`,
        `ppt/${this.targetType}s/_rels`,
      )) as XmlRelationshipHelper;
      await masterRels.assertRelatedContent(this.sourceArchive, info, assert);
    }
  }

  /**
   * Adds slide to presentation
   * @internal
   * @returns slide to presentation
   */
  async addToPresentation(): Promise<void> {
    const relId = await XmlHelper.getNextRelId(
      this.targetArchive,
      'ppt/_rels/presentation.xml.rels',
    );
    await this.appendToSlideRel(this.targetArchive, relId, this.targetNumber);

    if (this.targetType === 'slide') {
      await this.appendToSlideList(this.targetArchive, relId);
    } else if (this.targetType === 'slideMaster') {
      await this.appendToSlideMasterList(this.targetArchive, relId);
    } else if (this.targetType === 'slideLayout') {
      // No changes to ppt/presentation.xml required for slideLayouts
    }

    await this.appendToContentType(this.targetArchive, this.targetNumber);
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
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${this.targetType}`,
        Target: `${this.targetType}s/${this.targetType}${slideCount}.xml`,
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
   * Appends a new slide to slide list in presentation.xml.
   * If rootArchive has no slides, a new node will be created.
   * "id"-attribute of 'p:sldId'-element must be greater than 255.
   * @internal
   * @param rootArchive
   * @param relId
   * @returns to slide list
   */
  appendToSlideMasterList(
    rootArchive: IArchive,
    relId: string,
  ): Promise<HelperElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      parent: (xml: XmlDocument) =>
        xml.getElementsByTagName('p:sldMasterIdLst')[0],
      tag: 'p:sldMasterId',
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
  appendToContentType(
    rootArchive: IArchive,
    count: number,
  ): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/${this.targetType}s/${this.targetType}${count}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.${this.targetType}+xml`,
      }),
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
      await new Chart(
        {
          mode: 'append',
          target: chart,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceNumber,
        },
        this.targetType,
      ).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }

    const images = await Image.getAllOnSlide(this.sourceArchive, this.relsPath);
    for (const image of images) {
      await new Image(
        {
          mode: 'append',
          target: image,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceNumber,
        },
        this.targetType,
      ).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }
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
  async cleanSlide(targetPath: string): Promise<void> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetPath,
    );

    this.unsupportedTags.forEach((tag) => {
      const drop = xml.getElementsByTagName(tag);
      const length = drop.length;
      if (length && length > 0) {
        XmlHelper.sliceCollection(drop, 0);
      }
    });
    XmlHelper.writeXmlToArchive(this.targetArchive, targetPath, xml);
  }
}