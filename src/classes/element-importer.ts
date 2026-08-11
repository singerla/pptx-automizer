import { ElementNotFoundError } from '../errors';
import { log } from '../helper/logger';
import { GeneralHelper } from '../helper/general-helper';
import { PptPaths } from '../helper/ppt-paths';
import { analyzeElement } from '../helper/shape-type-detector';
import { XmlHelper } from '../helper/xml-helper';
import { XmlSlideHelper } from '../helper/xml-slide-helper';
import IArchive from '../interfaces/iarchive';
import { IShapeAction, ShapeActionMode } from '../interfaces/ishape-action';
import { ElementType } from '../enums/element-type';
import { Chart } from '../shapes/chart';
import { Diagram } from '../shapes/diagram';
import { GenericShape } from '../shapes/generic';
import { Hyperlink } from '../shapes/hyperlink';
import { Image } from '../shapes/image';
import { OLEObject } from '../shapes/ole';
import {
  ElementOnSlide,
  FindElementSelector,
  FindElementStrategy,
  ImportedElement,
  ImportElement,
  ShapeModificationCallback,
} from '../types/types';
import type HasShapes from './has-shapes';

/**
 * Owns the queue of elements to import/modify/remove on a slide, resolves
 * each entry to a source element (`getElementInfo`/`findElementOnSlide`)
 * and dispatches it to the matching shape class on write.
 *
 * Extracted from HasShapes (ROADMAP Phase 2).
 */
export class ElementImporter {
  /**
   * Queued element imports/modifications/removals of the owning slide.
   */
  queue: ImportElement[] = [];

  constructor(private host: HasShapes) {}

  /**
   * Queue an element action; executed on `automizer.write()`.
   */
  add(
    presName: string,
    slideNumber: number,
    selector: FindElementSelector,
    mode: ShapeActionMode,
    callback?: ShapeModificationCallback | ShapeModificationCallback[],
  ): void {
    this.queue.push({
      presName,
      slideNumber,
      selector,
      mode,
      callback,
    });
  }

  /**
   * Import all queued elements, merging multiple modifications
   * of the same source element.
   */
  async importSelected(): Promise<void> {
    const importElements = await this.getUniqueImportedElements();

    for (const element of importElements) {
      const info = element.info;
      const shape = info ? this.createShape(info) : undefined;
      if (!shape) {
        continue;
      }
      await this.runAction(shape, info.mode);
    }
  }

  /**
   * Processes and updates the list of imported elements by ensuring their
   * uniqueness based on a generated hash. If duplicate elements are found,
   * their callbacks are merged.
   */
  async getUniqueImportedElements(): Promise<ImportElement[]> {
    for (const element of this.queue) {
      const info = await this.getElementInfo(element);

      // getElementInfo only returns undefined with params.continueOnError;
      // skip the unresolvable element in that case.
      if (!info) {
        continue;
      }

      if (element.mode === 'append') {
        element.info = info;
        continue;
      }

      const selector = XmlSlideHelper.getSelector(info.sourceElement);
      const eleHash = JSON.stringify(selector);
      const alreadyImported = this.queue.find(
        (ele) => ele.info?.hash === eleHash,
      );
      if (alreadyImported) {
        const existingCallbacks = GeneralHelper.arrayify(element.callback);
        const pushCallbacks = GeneralHelper.arrayify(
          alreadyImported.info.callback,
        );
        alreadyImported.info.callback = [
          ...existingCallbacks,
          ...pushCallbacks,
        ];
      } else {
        info.hash = eleHash;
        element.info = info;
      }
    }

    return this.queue.filter((ele) => {
      return ele.info;
    });
  }

  /**
   * Instantiates the shape class matching the analyzed element type.
   * Returns undefined for elements that cannot be dispatched
   * (e.g. a hyperlink without a resolved target).
   */
  private createShape(info: ImportedElement): IShapeAction | undefined {
    const targetType = this.host.targetType;

    switch (info.type) {
      case ElementType.Chart:
        return new Chart(info, targetType);
      case ElementType.Image:
        return new Image(info, targetType);
      case ElementType.Shape:
        return new GenericShape(info, targetType);
      case ElementType.Diagram:
        return new Diagram(info, targetType);
      case ElementType.OLEObject:
        return new OLEObject(info, targetType, this.host.sourceArchive);
      case ElementType.Hyperlink:
        if (!info.target) {
          return undefined;
        }
        return new Hyperlink(
          info,
          targetType,
          this.host.sourceArchive,
          info.target.isExternal ? 'external' : 'internal',
          info.target.file,
        );
      default:
        return undefined;
    }
  }

  private async runAction(
    shape: IShapeAction,
    mode: ShapeActionMode,
  ): Promise<void> {
    const targetTemplate = this.host.targetTemplate;
    const targetNumber = this.host.targetNumber;

    switch (mode) {
      case 'append':
        await shape.append(targetTemplate, targetNumber);
        break;
      case 'modify':
        await shape.modify(targetTemplate, targetNumber);
        break;
      case 'remove':
        await shape.remove(targetTemplate, targetNumber);
        break;
    }
  }

  /**
   * Resolves a queued import to its source element and analyzed type.
   */
  async getElementInfo(importElement: ImportElement): Promise<ImportedElement> {
    const host = this.host;
    const template = host.root.getTemplate(importElement.presName);

    const slideNumber =
      importElement.mode === 'append'
        ? host.getSlideNumber(template, importElement.slideNumber)
        : importElement.slideNumber;

    let currentMode = 'slideToSlide';
    if (host.targetType === 'slideMaster') {
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
        ? PptPaths.slideMaster(slideNumber)
        : PptPaths.slide(slideNumber);

    const sourceRelPath =
      currentMode === 'onMaster'
        ? PptPaths.slideMasterRels(slideNumber)
        : PptPaths.slideRels(slideNumber);

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
      const message = `Can't find element on slide ${slideNumber} in ${importElement.presName} (selector: ${JSON.stringify(
        importElement.selector,
      )})`;
      if (host.presentation.params.continueOnError === true) {
        log.warn(message);
        return;
      }
      throw new ElementNotFoundError(message, {
        selector:
          typeof importElement.selector === 'string'
            ? importElement.selector
            : JSON.stringify(importElement.selector),
        file: sourcePath,
      });
    }

    const appendElementParams = await analyzeElement(
      sourceElement,
      sourceArchive,
      sourceRelPath,
    );

    return {
      mode: importElement.mode,
      name: selector,
      selector: XmlSlideHelper.getSelector(sourceElement),
      hasCreationId: mode === 'findByElementCreationId',
      sourceArchive,
      sourceSlideNumber: slideNumber,
      sourceElement,
      callback: importElement.callback,
      target: appendElementParams.target,
      type: appendElementParams.type,
      continueOnError: host.presentation.params.continueOnError === true,
    };
  }

  /**
   * Finds an element on a source slide by creationId and/or name.
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
    } else {
      if (selector.creationId) {
        strategies.push({
          mode: 'findByElementCreationId',
          selector: selector.creationId,
        });
      }

      strategies.push({
        mode: 'findByElementName',
        selector: selector.name,
        nameIdx: selector.nameIdx,
      });
    }

    for (const findElement of strategies) {
      const mode = findElement.mode;

      const sourceElement = await XmlHelper[mode](
        sourceArchive,
        sourcePath,
        findElement.selector,
        findElement.nameIdx,
      );

      if (sourceElement) {
        return { sourceElement, selector: findElement.selector, mode };
      }
    }

    return { sourceElement: undefined, selector: JSON.stringify(selector) };
  }
}
