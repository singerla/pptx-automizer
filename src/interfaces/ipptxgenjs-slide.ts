import PptxGenJS from 'pptxgenjs';

/**
 * Usage of pptxGenJs insinde pptx-automizer is restricted to supported
 * shape types.
 */
export interface IPptxGenJSSlide {
  /**
   * Add chart to Slide
   * @param {CHART_NAME|IChartMulti[]} type - chart type
   * @param {object[]} data - data object
   * @param {IChartOpts} options - chart options
   * @return {Slide} this Slide
   * @type {Function}
   */
  addChart(
    type: PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[],
    data: any[],
    options?: PptxGenJS.IChartOpts,
  ): void;

  /**
   * Add image to Slide
   * @param {ImageProps} options - image options
   * @return {Slide} this Slide
   */
  addImage(options: PptxGenJS.ImageProps): void;

  /**
   * Add shape to Slide
   * @param {SHAPE_NAME} shapeName - shape name
   * @param {ShapeProps} options - shape options
   * @return {Slide} this Slide
   */
  addShape(
    shapeName: PptxGenJS.SHAPE_NAME,
    options?: PptxGenJS.ShapeProps,
  ): void;

  /**
   * Add table to Slide
   * @param {TableRow[]} tableRows - table rows
   * @param {TableProps} options - table options
   * @return {Slide} this Slide
   */
  addTable(
    tableRows: PptxGenJS.TableRow[],
    options?: PptxGenJS.TableProps,
  ): void;

  /**
   * Add text to Slide
   * @param {string|TextProps[]} text - text string or complex object
   * @param {TextPropsOptions} options - text options
   * @return {Slide} this Slide
   */
  addText(
    text: string | PptxGenJS.TextProps[],
    options?: PptxGenJS.TextPropsOptions,
  ): void;
}
