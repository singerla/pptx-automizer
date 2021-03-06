export type ShapeCoordinates = {
  /**
   * Horizontal position measured from left corner
   * in PowerPoint `DXA` unit.
   * PowerPoint units are in `DXA` (except for font sizing)
   * - 1 inch is 1440 DXA
   * - 1 cm is 3658 DXA
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
