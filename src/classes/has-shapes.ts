import { SlideNotFoundError } from '../errors';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import IArchive from '../interfaces/iarchive';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import {
  AutomizerParams,
  FindElementSelector,
  GenerateElements,
  GenerateOnSlideCallback,
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
  ModifyXmlCallback,
  PlaceholderInfo,
} from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlSlideHelper } from '../helper/xml-slide-helper';
import { ContentTypeRegistry } from './content-type-registry';
import { ElementImporter } from './element-importer';
import { PlaceholderNormalizer } from './placeholder-normalizer';
import { RelatedContentCopier } from './related-content-copier';
import { SlideNotesCopier } from './slide-notes-copier';

/**
 * Base class of Slide, Master and Layout: holds the source/target context
 * of a copied part and the queues of deferred modifications.
 *
 * The actual work is done by collaborators (ROADMAP Phase 2):
 * - ElementImporter — queued element import/modify/remove
 * - RelatedContentCopier — charts/images/diagrams/OLE/hyperlinks
 * - SlideNotesCopier — notesSlide copy + number remapping
 * - PlaceholderNormalizer — placeholder cleanup, unsupported tags
 * - ContentTypeRegistry — presentation.xml + [Content_Types].xml entries
 */
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
   * Preparations of root template slide
   * @internal
   */
  preparations: SlideModificationCallback[];
  /**
   * Modifications of root template slide
   * @internal
   */
  modifications: SlideModificationCallback[];
  /**
   * Modifications of slide relations
   * @internal
   */
  relModifications: ModifyXmlCallback[];
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
    // exclude bullet images
    'a:buBlip',
    // 'p:oleObj',
    // 'mc:AlternateContent',
    // 'a14:imgProps',
    // 'a14:imgLayer'
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
  presentation: IPresentationProps;

  cleanupPlaceholders = false;

  /**
   * Collaborators doing the actual work on write (ROADMAP Phase 2).
   * @internal
   */
  elementImporter: ElementImporter;
  relatedContent: RelatedContentCopier;
  notes: SlideNotesCopier;
  placeholderNormalizer: PlaceholderNormalizer;
  contentTypes: ContentTypeRegistry;

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
  }) {
    this.sourceTemplate = params.template;

    this.preparations = [];
    this.modifications = [];
    this.relModifications = [];
    this.generateElements = [];

    this.presentation = params.presentation;

    this.status = params.presentation.status;
    this.content = params.presentation.content;

    this.cleanupPlaceholders = params.presentation.params.cleanupPlaceholders;

    this.elementImporter = new ElementImporter(this);
    this.relatedContent = new RelatedContentCopier(this);
    this.notes = new SlideNotesCopier(this);
    this.placeholderNormalizer = new PlaceholderNormalizer(this);
    this.contentTypes = new ContentTypeRegistry(this);
  }

  /**
   * Queued element imports/modifications/removals of this slide.
   * @internal
   */
  get importElements(): ImportElement[] {
    return this.elementImporter.queue;
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
   * @param filterTags Use an array of strings to filter parent tags (e.g. 'sp')
   * @param layoutPlaceholders
   * @returns {Promise<ElementInfo[]>} A promise that resolves to an array of ElementInfo objects.
   */
  async getAllElements(
    filterTags?: string[],
    layoutPlaceholders?: PlaceholderInfo[],
  ): Promise<ElementInfo[]> {
    const xmlSlideHelper = await this.getSlideHelper();

    // Get all ElementInfo objects
    return xmlSlideHelper.getAllElements(filterTags, layoutPlaceholders);
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
    return this.getSlideHelperInstance(
      this.sourceTemplate.archive,
      this.sourcePath,
      this.sourceNumber,
    );
  }

  async getSlideHelperInstance(
    archive: IArchive,
    path: string,
    number: number,
  ): Promise<XmlSlideHelper> {
    try {
      // Retrieve the slide XML data
      const slideXml = await XmlHelper.getXmlFromArchive(archive, path);

      const sourceLayoutId = await XmlRelationshipHelper.getSlideLayoutNumber(
        archive,
        number,
      );

      // Initialize the XmlSlideHelper
      return new XmlSlideHelper(slideXml, {
        sourceArchive: archive,
        slideNumber: number,
        sourceLayoutId,
      });
    } catch (error) {
      // Log the error message
      throw new Error(error.message);
    }
  }

  /**
   * Push preparations list
   * @internal
   * @param callback
   */
  prepare(callback: SlideModificationCallback): void {
    this.preparations.push(callback);
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
  modifyRelations(callback: ModifyXmlCallback): void {
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
    this.elementImporter.add(
      this.sourceTemplate.name,
      this.sourceNumber,
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
    this.elementImporter.add(
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
    this.elementImporter.add(
      this.sourceTemplate.name,
      this.sourceNumber,
      selector,
      'remove',
      undefined,
    );

    return this;
  }

  /**
   * ToDo: Implement creationIds as well for slideMasters
   *
   * Try to convert a given slide's creationId to corresponding slide number.
   * Used if automizer is run with useCreationIds: true
   * @internal
   * @param template
   * @param slideIdentifier
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

      throw new SlideNotFoundError(
        'Could not find slide number for creationId: ' +
          slideIdentifier +
          '@' +
          template.name,
        { slideIdentifier, templateName: template.name },
      );
    }

    return slideIdentifier as number;
  }

  /**
   * Imported selected elements while merging multiple element modifications
   * @internal
   */
  async importedSelectedElements(): Promise<void> {
    await this.elementImporter.importSelected();
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
    await this.contentTypes.addToPresentation();
  }

  /**
   * Copys related content
   * @internal
   * @returns related content
   */
  async copyRelatedContent(): Promise<void> {
    await this.relatedContent.copy();
  }

  /**
   * Applys slide preparation callbacks
   * Will be executed before any shape modifications callback
   * @internal
   * @returns modifications
   */
  async applyPreparations(): Promise<void> {
    for (const modification of this.preparations) {
      const xml = await XmlHelper.getXmlFromArchive(
        this.targetArchive,
        this.targetPath,
      );
      await modification(xml, this);
      XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, xml);
    }
  }

  /**
   * Applys slide modification callbacks
   * Will be executed after all shape modifications callbacks
   * @internal
   * @returns modifications
   */
  async applyModifications(): Promise<void> {
    for (const modification of this.modifications) {
      const xml = await XmlHelper.getXmlFromArchive(
        this.targetArchive,
        this.targetPath,
      );
      await modification(xml, this);
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
      PptPaths.partRels(this.targetType, this.targetNumber),
      this.relModifications,
    );
  }

  /**
   * Removes all unsupported tags from slide xml and (optionally)
   * normalizes placeholder shapes.
   * @internal
   */
  async cleanSlide(
    targetPath: string,
    sourcePlaceholderTypes?: SlidePlaceholder[],
  ): Promise<void> {
    await this.placeholderNormalizer.cleanSlide(
      targetPath,
      sourcePlaceholderTypes,
    );
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
    return this.placeholderNormalizer.parsePlaceholders();
  }

  /**
   * Flushes this part's finished target XML (and its rels) out of the
   * archive's DOM buffer. Called as the last step of append(): keeping every
   * appended slide's parsed DOM alive made memory grow with deck size
   * (~8 MB per large slide). Anything reading the part afterwards
   * (e.g. `cleanup` at write time) re-parses it from the archive.
   * @internal
   */
  async flushTargetXml(): Promise<void> {
    await this.targetArchive.flushXml(this.targetPath);
    await this.targetArchive.flushXml(this.targetRelsPath);
  }
}
