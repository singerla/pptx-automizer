import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import IArchive from '../interfaces/iarchive';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import {
  AnalyzedElementType,
  AutomizerParams,
  ElementOnSlide,
  FindElementSelector,
  FindElementStrategy,
  GenerateElements,
  GenerateOnSlideCallback,
  ImportedElement,
  ImportElement,
  ShapeModificationCallback,
  ShapeTargetType,
  SlideModificationCallback,
  SlidePlaceholder,
  SourceIdentifier,
  StatusTracker,
} from '../types/types';
import { ContentTracker } from '../helper/content-tracker';
import {
  ElementInfo,
  RelationshipAttribute,
  SlideListAttribute,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { FileHelper } from '../helper/file-helper';
import { Chart } from '../shapes/chart';
import { Image } from '../shapes/image';
import { ElementType } from '../enums/element-type';
import { GenericShape } from '../shapes/generic';
import { XmlSlideHelper } from '../helper/xml-slide-helper';
import { OLEObject } from '../shapes/ole';
import { Hyperlink } from '../shapes/hyperlink';

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
   * Modifications of root template slide
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
   * Generate elements on slide with PptxGenJS
   * @internal
   */
  generateElements: GenerateElements[];
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
    // 'p:oleObj',
    // 'mc:AlternateContent',
    //'a14:imgProps',
  ];
  /**
   * List of unsupported tags in slide xml
   * @internal
   */
  unsupportedRelationTypes = [
    //  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags',
  ];
  targetType: ShapeTargetType;
  params: AutomizerParams;

  cleanupPlaceholders: boolean = false;

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
  }) {
    this.sourceTemplate = params.template;

    this.modifications = [];
    this.relModifications = [];
    this.importElements = [];
    this.generateElements = [];

    this.status = params.presentation.status;
    this.content = params.presentation.content;

    this.cleanupPlaceholders = params.presentation.params.cleanupPlaceholders;
  }

  /**
   * Asynchronously retrieves all text element IDs from the slide.
   * @returns {Promise<string[]>} A promise that resolves to an array of text element IDs.
   */
  async getAllTextElementIds(): Promise<string[]> {
    const xmlSlideHelper = await this.getSlideHelper();

    // Get all text element IDs
    return xmlSlideHelper.getAllTextElementIds(
      this.sourceTemplate.useCreationIds || false,
    );
  }

  /**
   * Asynchronously retrieves all elements from the slide.
   * @params filterTags Use an array of strings to filter parent tags (e.g. 'sp')
   * @returns {Promise<ElementInfo[]>} A promise that resolves to an array of ElementInfo objects.
   */
  async getAllElements(filterTags?: string[]): Promise<ElementInfo[]> {
    const xmlSlideHelper = await this.getSlideHelper();

    // Get all ElementInfo objects
    return xmlSlideHelper.getAllElements(filterTags);
  }

  /**
   * Asynchronously retrieves one element from the slide.
   * @params selector Use shape name or creationId to find the shape
   * @returns {Promise<ElementInfo>} A promise that resolves an ElementInfo object.
   */
  async getElement(selector: string): Promise<ElementInfo> {
    const xmlSlideHelper = await this.getSlideHelper();
    return xmlSlideHelper.getElement(selector);
  }

  /**
   * Asynchronously retrieves the dimensions of the slide.
   * This function utilizes the XmlSlideHelper to get the slide dimensions.
   *
   * @returns {Promise<{width: number, height: number}>} A promise that resolves to an object containing the width and height of the slide.
   */
  async getDimensions(): Promise<{ width: number; height: number }> {
    const xmlSlideHelper = await this.getSlideHelper();
    return xmlSlideHelper.getDimensions();
  }

  /**
   * Asynchronously retrieves an instance of XmlSlideHelper for slide.
   * @returns {Promise<XmlSlideHelper>} An instance of XmlSlideHelper.
   */
  async getSlideHelper(): Promise<XmlSlideHelper> {
    try {
      // Retrieve the slide XML data
      const slideXml = await XmlHelper.getXmlFromArchive(
        this.sourceTemplate.archive,
        this.sourcePath,
      );

      // Initialize the XmlSlideHelper
      return new XmlSlideHelper(slideXml, this);
    } catch (error) {
      // Log the error message
      throw new Error(error.message);
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
   * Push relations modifications list
   * @internal
   * @param callback
   */
  modifyRelations(callback: SlideModificationCallback): void {
    this.relModifications.push(callback);
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

    this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'modify',
      callback,
    );

    return this;
  }

  generate(generate: GenerateOnSlideCallback, objectName?: string): this {
    this.generateElements.push({
      objectName,
      callback: generate,
    });
    return this;
  }

  getGeneratedElements(): GenerateElements[] {
    return this.generateElements;
  }

  /**
   * Select, insert and (optionally) modify a single element to a slide.
   * @param {string} presName - Filename or alias name of the template presentation.
   * Must have been importet with Automizer.load().
   * @param {number} slideNumber - Slide number within the specified template to search for the required element.
   * @param {FindElementSelector} selector - a string or object to find the target element
   * @param {ShapeModificationCallback | ShapeModificationCallback[]} callback - One or more callback functions to apply.
   * Depending on the shape type (e.g. chart or table), different arguments will be passed to the callback.
   */
  addElement(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): this {
    this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'append',
      callback,
    );

    return this;
  }

  /**
   * Remove a single element from slide.
   * @param {string} selector - Element's name on the slide.
   */
  removeElement(selector: FindElementSelector): this {
    const presName = this.sourceTemplate.name;
    const slideNumber = this.sourceNumber;

    this.addElementToModificationsList(
      presName,
      slideNumber,
      selector,
      'remove',
      undefined,
    );

    return this;
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
  ): void {
    this.importElements.push({
      presName,
      slideNumber,
      selector,
      mode,
      callback,
    });
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
        case ElementType.OLEObject:
          await new OLEObject(info, this.targetType, this.sourceArchive)[
            info.mode
          ](this.targetTemplate, this.targetNumber, this.targetType);
          break;
        case ElementType.Hyperlink:
          // For hyperlinks, we need to handle them differently
          if (info.target) {
            await new Hyperlink(
              info,
              this.targetType,
              this.sourceArchive,
              info.target.isExternal ? 'external' : 'internal',
              info.target.file,
            )[info.mode](this.targetTemplate, this.targetNumber);
          }
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

    let currentMode = 'slideToSlide';
    if (this.targetType === 'slideMaster') {
      if (importElement.mode === 'append') {
        currentMode = 'slideToMaster';
      } else {
        currentMode = 'onMaster';
      }
    }

    // It is possible to import shapes from loaded slides to slideMaster,
    // as well as to modify an existing shape on current slideMaster
    const sourcePath =
      currentMode === 'onMaster'
        ? `ppt/slideMasters/slideMaster${slideNumber}.xml`
        : `ppt/slides/slide${slideNumber}.xml`;

    const sourceRelPath =
      currentMode === 'onMaster'
        ? `ppt/slideMasters/_rels/slideMaster${slideNumber}.xml.rels`
        : `ppt/slides/_rels/slide${slideNumber}.xml.rels`;

    const sourceArchive = await template.archive;
    const useCreationIds =
      template.useCreationIds === true && template.creationIds !== undefined;

    const { sourceElement, selector, mode } = await this.findElementOnSlide(
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
      sourceRelPath,
    );

    return {
      mode: importElement.mode,
      name: selector,
      hasCreationId: mode === 'findByElementCreationId',
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
  ): Promise<ElementOnSlide> {
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
        return { sourceElement, selector: findElement.selector, mode };
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
  ): Promise<XmlElement> {
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
  appendToSlideList(rootArchive: IArchive, relId: string): Promise<XmlElement> {
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
  ): Promise<XmlElement> {
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
  ): Promise<XmlElement> {
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
  ): Promise<XmlElement> {
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

    const oleObjects = await OLEObject.getAllOnSlide(
      this.sourceArchive,
      this.relsPath,
    );
    for (const oleObject of oleObjects) {
      await new OLEObject(
        {
          mode: 'append',
          target: oleObject,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceNumber,
        },
        this.targetType,
        this.sourceArchive,
      ).modifyOnAddedSlide(this.targetTemplate, this.targetNumber, oleObjects);
    }

    // Copy hyperlinks
    const hyperlinks = await Hyperlink.getAllOnSlide(
      this.sourceArchive,
      this.relsPath,
    );
    for (const hyperlink of hyperlinks) {
      // Create a new hyperlink with the correct target information
      const hyperlinkInstance = new Hyperlink(
        {
          mode: 'append',
          target: hyperlink,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceNumber,
          sourceRid: hyperlink.rId,
        },
        this.targetType,
        this.sourceArchive,
        hyperlink.isExternal ? 'external' : 'internal',
        hyperlink.file,
      );

      // Ensure the target property is properly set
      hyperlinkInstance.target = hyperlink;

      // Process the hyperlink
      await hyperlinkInstance.modifyOnAddedSlide(
        this.targetTemplate,
        this.targetNumber,
        hyperlinks,
      );
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
    sourceElement: XmlElement,
    sourceArchive: IArchive,
    relsPath: string,
  ): Promise<AnalyzedElementType> {
    const isChart = sourceElement.getElementsByTagName('c:chart');

    if (isChart.length) {
      const target = await XmlHelper.getTargetByRelId(
        sourceArchive,
        relsPath,
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
        relsPath,
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
          relsPath,
          sourceElement,
          'image',
        ),
      } as AnalyzedElementType;
    }

    const isOLEObject = sourceElement.getElementsByTagName('p:oleObj');
    if (isOLEObject.length) {
      const target = await XmlHelper.getTargetByRelId(
        sourceArchive,
        relsPath,
        sourceElement,
        'oleObject',
      );

      return {
        type: ElementType.OLEObject,
        target: target,
      } as AnalyzedElementType;
    }

    // Check for hyperlinks in text runs
    const hasHyperlink = this.findHyperlinkInElement(sourceElement);
    if (hasHyperlink) {
      try {
        const target = await XmlHelper.getTargetByRelId(
          sourceArchive,
          relsPath,
          sourceElement,
          'hyperlink',
        );

        return {
          type: ElementType.Hyperlink,
          target: target,
          element: sourceElement,
        } as AnalyzedElementType;
      } catch (error) {
        console.warn('Error finding hyperlink target:', error);
      }
    }

    return {
      type: ElementType.Shape,
    } as AnalyzedElementType;
  }

  // Helper method to find hyperlinks in an element
  private findHyperlinkInElement(element: XmlElement): boolean {
    // Check for direct hyperlinks
    const directHyperlinks = element.getElementsByTagName('a:hlinkClick');
    if (directHyperlinks.length > 0) {
      return true;
    }

    // Check for hyperlinks in text runs
    const textRuns = element.getElementsByTagName('a:r');
    for (let i = 0; i < textRuns.length; i++) {
      const run = textRuns[i];
      const rPr = run.getElementsByTagName('a:rPr')[0];
      if (rPr && rPr.getElementsByTagName('a:hlinkClick').length > 0) {
        return true;
      }
    }

    return false;
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
   * Apply modifications to slide relations
   * @internal
   * @returns modifications
   */
  async applyRelModifications(): Promise<void> {
    await XmlHelper.modifyXmlInArchive(
      this.targetArchive,
      `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetNumber}.xml.rels`,
      this.relModifications,
    );
  }

  /**
   * Removes all unsupported tags from slide xml.
   * E.g. added relations & tags by Thinkcell cannot
   * be processed by pptx-automizer at the moment.
   * @internal
   */
  async cleanSlide(
    targetPath: string,
    sourcePlaceholderTypes?: SlidePlaceholder[],
  ): Promise<void> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetPath,
    );

    if (this.cleanupPlaceholders && sourcePlaceholderTypes) {
      this.removeDuplicatePlaceholders(xml, sourcePlaceholderTypes);
      this.normalizePlaceholderShapes(xml, sourcePlaceholderTypes);
    }

    this.unsupportedTags.forEach((tag) => {
      const drop = xml.getElementsByTagName(tag);
      const length = drop.length;
      if (length && length > 0) {
        XmlHelper.sliceCollection(drop, 0);
      }
    });
    XmlHelper.writeXmlToArchive(this.targetArchive, targetPath, xml);
  }

  /**
   * If you insert a placeholder shape on a target slide with an empty
   * placeholder of the same type, we need to remove the existing
   * placeholder.
   *
   * @param xml
   * @param sourcePlaceholderTypes
   */
  removeDuplicatePlaceholders(
    xml: XmlDocument,
    sourcePlaceholderTypes: SlidePlaceholder[],
  ) {
    const placeholders = xml.getElementsByTagName('p:ph');
    const usedTypes = {};
    XmlHelper.modifyCollection(placeholders, (placeholder: XmlElement) => {
      const type = placeholder.getAttribute('type');
      usedTypes[type] = usedTypes[type] || 0;
      usedTypes[type]++;
    });

    for (const usedType in usedTypes) {
      const count = usedTypes[usedType];
      if (count > 1) {
        // TODO: in case more than two placeholders are of a kind,
        // this will likely remove more than intended. Should also match by id.
        const removePlaceholders = sourcePlaceholderTypes.filter(
          (sourcePlaceholder) => sourcePlaceholder.type === usedType,
        );
        removePlaceholders.forEach((removePlaceholder) => {
          const parentShapeTag = 'p:sp';
          const parentShape = XmlHelper.getClosestParent(
            parentShapeTag,
            removePlaceholder.xml,
          );
          if (parentShape) {
            XmlHelper.remove(parentShape);
          }
        });
      }
    }
  }

  /**
   * If a placeholder shape was inserted on a slide without a corresponding
   * placeholder, powerPoint will usually smash the shape's formatting.
   * This function removes the placeholder tag.
   * @param xml
   * @param sourcePlaceholderTypes
   */
  normalizePlaceholderShapes(
    xml: XmlDocument,
    sourcePlaceholderTypes: SlidePlaceholder[],
  ) {
    const placeholders = xml.getElementsByTagName('p:ph');
    XmlHelper.modifyCollection(placeholders, (placeholder: XmlElement) => {
      const usedType = placeholder.getAttribute('type');
      const existingPlaceholder = sourcePlaceholderTypes.find(
        (sourcePlaceholder) => sourcePlaceholder.type === usedType,
      );
      if (!existingPlaceholder) {
        XmlHelper.remove(placeholder);
      }
    });
  }

  /**
   * Removes all unsupported relations from _rels xml.
   * @internal
   */
  async cleanRelations(targetRelsPath: string): Promise<void> {
    await XmlHelper.removeIf({
      archive: this.targetArchive,
      file: targetRelsPath,
      tag: 'Relationship',
      clause: (xml, item) => {
        return this.unsupportedRelationTypes.includes(
          item.getAttribute('Type'),
        );
      },
    });
  }

  async parsePlaceholders(): Promise<SlidePlaceholder[]> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetPath,
    );
    const placeholderTypes = [];
    const placeholders = xml.getElementsByTagName('p:ph');
    XmlHelper.modifyCollection(placeholders, (placeholder: XmlElement) => {
      placeholderTypes.push({
        type: placeholder.getAttribute('type'),
        id: placeholder.getAttribute('id'),
        xml: placeholder,
      });
    });
    return placeholderTypes;
  }
}
