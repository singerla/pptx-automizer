import { RootPresTemplate } from './root-pres-template';

/**
 * The three things that can happen to a queued element on write.
 */
export type ShapeActionMode = 'append' | 'modify' | 'remove';

/**
 * Implemented by all importable shape classes (Chart, Image, GenericShape,
 * Diagram, OLEObject, Hyperlink). Replaces the former string-indexed
 * dispatch (`new Chart(info)[info.mode](…)`) with type-checked calls.
 *
 * A shape type that does not support an action implements the method by
 * throwing a descriptive error instead of omitting it.
 */
export interface IShapeAction {
  append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<unknown>;
  modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<unknown>;
  remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<unknown>;
}
