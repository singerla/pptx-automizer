import { Color } from './modify-types';

export type ShapeCoordinates = {
  /**
   * Horizontal position measured from left corner
   * in PowerPoint `DXA` unit.
   * PowerPoint units are in `DXA` (except for font sizing)
   * - 1 cm is 360000 DXA
   * - 20 DXA is 1 point
   */
  x?: number;
  /**
   * Vertical position on the slide
   * Measured from top corner in `DXA`
   */
  y?: number;
  /**
   * Shape width in `DXA`
   */
  w?: number;
  /**
   * Shape height in `DXA`
   */
  h?: number;
};
export type ShapeValueStyle = {
  background?: Color;
  gradient?: {
    color: Color;
    index: number;
  }[];
  border?: {
    color?: Color;
    weight?: number;
  };
  label?: {
    color?: Color;
    isBold?: boolean;
    size?: number;
  };
};
