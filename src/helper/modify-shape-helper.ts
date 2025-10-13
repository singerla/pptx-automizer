/**
 * @file ModifyShapeHelper provides utility functions for manipulating PowerPoint shapes
 * through XML modifications.
 */
import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import { GeneralHelper, vd } from './general-helper';
import TextReplaceHelper from './text-replace-helper';
import ModifyTextHelper from './modify-text-helper';
import { XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { XmlSlideHelper } from './xml-slide-helper';
import { ModifyColorHelper } from '../index';

/**
 * Mapping between user-friendly property names and their corresponding XML structure
 * This allows using multiple alias names for the same shape properties
 */
const map = {
  // X position mappings (left)
  x: { tag: 'a:off', attribute: 'x' },
  l: { tag: 'a:off', attribute: 'x' },
  left: { tag: 'a:off', attribute: 'x' },

  // Y position mappings (top)
  y: { tag: 'a:off', attribute: 'y' },
  t: { tag: 'a:off', attribute: 'y' },
  top: { tag: 'a:off', attribute: 'y' },

  // Width mappings
  cx: { tag: 'a:ext', attribute: 'cx' },
  w: { tag: 'a:ext', attribute: 'cx' },
  width: { tag: 'a:ext', attribute: 'cx' },

  // Height mappings
  cy: { tag: 'a:ext', attribute: 'cy' },
  h: { tag: 'a:ext', attribute: 'cy' },
  height: { tag: 'a:ext', attribute: 'cy' },
};

/**
 * Helper class for modifying PowerPoint shapes in XML structure
 * Provides various methods for manipulating shape appearance, position, and content
 */
export default class ModifyShapeHelper {
  /**
   * Set solid fill of modified shape to accent6 color
   *
   * @param element - XML element representing the shape
   */
  static setSolidFill = (element: XmlElement): void => {
    element
      .getElementsByTagName('a:solidFill')[0]
      .getElementsByTagName('a:schemeClr')[0]
      .setAttribute('val', 'accent6');
  };

  /**
   * Set text content of a shape
   *
   * @param text - Text string to set as shape content
   * @returns Function that accepts an XML element to modify
   */
  static setText =
    (text: string) =>
    (element: XmlElement): void => {
      ModifyTextHelper.setText(text)(element as XmlElement);
    };

  /**
   * Set content to bulleted list within a shape
   *
   * @param list - Array or string content for the bullet list
   * @returns Function that accepts an XML element to modify
   */
  static setBulletList =
    (list) =>
    (element: XmlElement): void => {
      ModifyTextHelper.setBulletList(list)(element as XmlElement);
    };

  /**
   * Replace tagged text content within modified shape
   *
   * @param replaceText - Single replacement or array of replacements
   * @param options - Optional configuration for text replacement
   * @returns Function that accepts an XML element to modify
   */
  static replaceText =
    (replaceText: ReplaceText | ReplaceText[], options?: ReplaceTextOptions) =>
    (element: XmlElement): void => {
      const replaceTexts = GeneralHelper.arrayify(replaceText);

      new TextReplaceHelper(options, element as XmlElement)
        .isolateTaggedNodes()
        .applyReplacements(replaceTexts);
    };

  /**
   * Creates missing transformation elements (a:off and a:ext) with default values
   * Ensures that a shape has the required XML structure for positioning and sizing
   *
   * @param element - The XML element to check and modify
   * @returns The xfrm element that contains or will contain the transformation data
   */
  static ensureTransformElements = (element: XmlElement): XmlElement | null => {
    // First find the xfrm element (could be a:xfrm or p:xfrm)
    let xfrm = element.getElementsByTagName('a:xfrm')[0] as XmlElement;
    if (!xfrm) {
      xfrm = element.getElementsByTagName('p:xfrm')[0] as XmlElement;
    }

    // If no xfrm element exists, try to find the appropriate parent to create it
    if (!xfrm) {
      const spPr =
        element.getElementsByTagName('p:spPr')[0] ||
        element.getElementsByTagName('a:spPr')[0];

      if (!spPr) {
        return null; // Cannot create xfrm without a proper parent
      }

      // Create the xfrm element
      xfrm = element.ownerDocument.createElement('a:xfrm');
      spPr.appendChild(xfrm);
    }

    // Create a:off element if it doesn't exist
    if (!xfrm.getElementsByTagName('a:off')[0]) {
      const newAOff = element.ownerDocument.createElement('a:off');
      // Set default coordinates (0,0)
      newAOff.setAttribute('x', '0');
      newAOff.setAttribute('y', '0');

      // Insert a:off as the first child of xfrm
      if (xfrm.hasChildNodes()) {
        xfrm.insertBefore(newAOff, xfrm.firstChild);
      } else {
        xfrm.appendChild(newAOff);
      }
    }

    // Create a:ext element if it doesn't exist
    if (!xfrm.getElementsByTagName('a:ext')[0]) {
      const newAExt = element.ownerDocument.createElement('a:ext');
      // Set default dimensions - using 1000000 which is about 2.78cm
      newAExt.setAttribute('cx', '1000000');
      newAExt.setAttribute('cy', '1000000');

      // Add the a:ext element after a:off
      xfrm.appendChild(newAExt);
    }

    return xfrm;
  };

  /**
   * Set absolute position and size of a shape
   *
   * @param pos - Object containing position and size coordinates
   * @returns Function that accepts an XML element to modify
   */
  static setPosition =
    (pos: ShapeCoordinates) =>
    (element: XmlElement): void => {
      // Ensure the transform elements exist
      const xfrm = ModifyShapeHelper.ensureTransformElements(element);
      if (!xfrm) {
        return; // Cannot set position without proper structure
      }

      // Apply each provided coordinate
      Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        // Skip invalid values or unsupported properties
        if (typeof value !== 'number' || !map[key]) return;
        // Ensure value is not negative
        value = value < 0 ? 0 : value;

        // Set the value in the appropriate XML tag and attribute
        xfrm
          .getElementsByTagName(map[key].tag)[0]
          .setAttribute(map[key].attribute, value);
      });
    };

  /**
   * Incrementally update position and size of a shape by adding delta values
   *
   * @param pos - Object containing delta values for position and size
   * @returns Function that accepts an XML element to modify
   */
  static updatePosition =
    (pos: ShapeCoordinates) =>
    (element: XmlElement): void => {
      // Ensure the transform elements exist
      const xfrm = ModifyShapeHelper.ensureTransformElements(element);
      if (!xfrm) {
        return; // Cannot update position without proper structure
      }

      // Apply each provided delta coordinate
      Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        // Skip invalid values or unsupported properties
        if (typeof value !== 'number' || !map[key]) return;

        // Get current value and add the delta
        const currentValue = xfrm
          .getElementsByTagName(map[key].tag)[0]
          .getAttribute(map[key].attribute);

        value += Number(currentValue);

        // Update the value in the appropriate XML tag and attribute
        xfrm
          .getElementsByTagName(map[key].tag)[0]
          .setAttribute(map[key].attribute, value);
      });
    };

  /**
   * Rotate a shape by a given angle in degrees
   *
   * @param degrees - Rotation angle in degrees (positive = clockwise, negative = counterclockwise)
   * @returns Function that accepts an XML element to modify
   */
  static rotate =
    (degrees: number) =>
    (element: XmlElement): void => {
      const spPr = element.getElementsByTagName('p:spPr');

      if (spPr) {
        const xfrm = spPr.item(0).getElementsByTagName('a:xfrm').item(0);
        // Convert negative degrees to equivalent positive value (0-359)
        degrees = degrees < 0 ? 360 + degrees : degrees;
        // PowerPoint uses 60000 units per degree for rotation
        xfrm.setAttribute('rot', String(Math.round(degrees * 60000)));
      }
    };

  /**
   * Apply rounded corners to a shape with a specified corner radius
   *
   * @param degree - Corner radius in EMU units (1 cm = 360000 EMU)
   * @returns Function that accepts an XML element to modify
   */
  static roundedCorners =
    (degree: number) =>
    (element: XmlElement): void => {
      // Find the spPr element where we need to add or modify the a:prstGeom element
      const spPr =
        element.getElementsByTagName('p:spPr')[0] ||
        element.getElementsByTagName('a:spPr')[0];

      if (!spPr) {
        return; // Cannot find spPr element
      }

      // Get the shape dimensions to calculate the appropriate adjustment value
      const xfrm = ModifyShapeHelper.ensureTransformElements(element);
      if (!xfrm) {
        return; // Cannot proceed without proper transformation data
      }

      // Get current width and height
      const width = Number(
        xfrm.getElementsByTagName('a:ext')[0].getAttribute('cx'),
      );
      const height = Number(
        xfrm.getElementsByTagName('a:ext')[0].getAttribute('cy'),
      );

      // Calculate the adjustment value based on the smaller dimension
      const minDimension = Math.min(width, height);

      // Ensure degree is within reasonable bounds (PowerPoint uses 0-50% for rounded rect)
      const clampedDegree = Math.max(0, Math.min(degree, minDimension / 2));

      // Calculate the adjustment value (0-100000 where 100000 is 100%)
      // PowerPoint uses the percentage of the shorter dimension for corners
      const adjValue = Math.round((clampedDegree / minDimension) * 100000);

      // Remove any existing prstGeom element
      const existingPrstGeom = spPr.getElementsByTagName('a:prstGeom')[0];
      if (existingPrstGeom) {
        XmlHelper.remove(existingPrstGeom)
      }

      // Create the new prstGeom element with the roundRect preset
      const prstGeom = element.ownerDocument.createElement('a:prstGeom');
      prstGeom.setAttribute('prst', 'roundRect');

      // Create the avLst element and the adjustment value
      const avLst = element.ownerDocument.createElement('a:avLst');
      const gd = element.ownerDocument.createElement('a:gd');
      gd.setAttribute('name', 'adj');
      gd.setAttribute('fmla', `val ${adjValue}`);

      // Build the element hierarchy
      avLst.appendChild(gd);
      prstGeom.appendChild(avLst);

      // Add the new prstGeom element to spPr at the appropriate position
      // It should be added after xfrm but before other elements
      const xfrmElement = spPr.getElementsByTagName('a:xfrm')[0];

      if (xfrmElement && xfrmElement.nextSibling) {
        spPr.insertBefore(prstGeom, xfrmElement.nextSibling);
      } else {
        spPr.appendChild(prstGeom);
      }

      // Check for noFill element - for picture elements, we need to remove noFill
      // otherwise the picture will be invisible when we apply rounded corners
      const noFillElement = spPr.getElementsByTagName('a:noFill')[0];
      if (noFillElement) {
        XmlHelper.remove(noFillElement)
      }
    };

  /**
   * Removes background color and fill elements from a shape
   * This removes both visible and hidden fill properties and sets explicit noFill
   *
   * @param element - The XML element representing the shape
   */
  static removeBackground = (element: XmlElement): void => {
    // Find the spPr (ShapeProperties) element where fill properties are defined
    const spPr =
      element.getElementsByTagName('p:spPr')[0] ||
      element.getElementsByTagName('a:spPr')[0];

    if (!spPr) {
      return; // No shape properties found, nothing to remove
    }

    // Remove all types of fill elements
    // 1. solidFill - used for solid color backgrounds
    const solidFill = spPr.getElementsByTagName('a:solidFill')[0];
    if (solidFill) {
      XmlHelper.remove(solidFill);
    }

    // 2. gradFill - used for gradient backgrounds
    const gradFill = spPr.getElementsByTagName('a:gradFill')[0];
    if (gradFill) {
      XmlHelper.remove(gradFill);
    }

    // 3. pattFill - used for pattern backgrounds
    const pattFill = spPr.getElementsByTagName('a:pattFill')[0];
    if (pattFill) {
      XmlHelper.remove(pattFill);
    }

    // 4. grpFill - used when inheriting fill from parent group
    const grpFill = spPr.getElementsByTagName('a:grpFill')[0];
    if (grpFill) {
      XmlHelper.remove(grpFill);
    }
  };

  /**
   * Removes shape style but preserves fill color information by moving it to a standard fill element
   * This converts the p:style reference colors to direct solidFill colors on the shape
   *
   * @param element - The XML element representing the shape
   */
  static removeShapeStyle = (element: XmlElement): void => {
    // Find the style element in the shape
    const styleElement = element.getElementsByTagName('p:style')[0];
    if (!styleElement) {
      return; // No style element found, nothing to do
    }

    // First extract the fill color information before removing the style
    const fillRef = styleElement.getElementsByTagName('a:fillRef')[0];
    let colorInfo = null;

    if (fillRef) {
      // Try to get the scheme color information
      const schemeClr = fillRef.getElementsByTagName('a:schemeClr')[0];
      if (schemeClr) {
        const colorValue = schemeClr.getAttribute('val');
        if (colorValue) {
          colorInfo = {
            type: 'schemeClr',
            value: colorValue,
          };
        }
      }

      // Check for srgbClr as an alternative
      const srgbClr = fillRef.getElementsByTagName('a:srgbClr')[0];
      if (srgbClr && !colorInfo) {
        const colorValue = srgbClr.getAttribute('val');
        if (colorValue) {
          colorInfo = {
            type: 'srgbClr',
            value: colorValue,
          };
        }
      }
    }

    // Now remove the style element
    XmlHelper.remove(styleElement);

    // If we found color information, apply it to the shape as a standard fill
    if (colorInfo) {
      ModifyColorHelper.solidFill(colorInfo)(
        element.getElementsByTagName('p:spPr').item(0),
      );
    }
  };

  /**
   * Removes border/outline from a shape
   * This function removes the a:ln (line) element which defines the border properties
   *
   * @param element - The XML element representing the shape
   */
  static removeBorder = (element: XmlElement): void => {
    // Find the spPr (ShapeProperties) element where line/border properties are defined
    const spPr =
      element.getElementsByTagName('p:spPr')[0] ||
      element.getElementsByTagName('a:spPr')[0];

    if (!spPr) {
      return; // No shape properties found, nothing to remove
    }

    // Remove the a:ln element which defines the border/outline
    const ln = spPr.getElementsByTagName('a:ln')[0];
    if (ln) {
      XmlHelper.remove(ln);
    }

    // Also check for table cell borders (lnL, lnR, lnT, lnB) if they exist
    // These are used for table cell borders in PowerPoint
    const tableBorders = ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'];
    tableBorders.forEach((borderTag) => {
      const border = element.getElementsByTagName(borderTag)[0];
      if (border) {
        XmlHelper.remove(border);
      }
    });
  };

  /**
   * Removes unwanted visual effects from PowerPoint elements based on their type
   * Preserves certain effects for drawings but removes distracting effects from photos
   *
   * @param element - The XML element to clean up
   */
  static removeImageEffects(element: XmlElement): void {
    // Determine the visual type of the element (picture, chart, etc.)
    const elementType = XmlSlideHelper.getElementVisualType(element);

    // Common elements to process
    const spPr =
      element.getElementsByTagName('p:spPr')[0] ||
      element.getElementsByTagName('a:spPr')[0];

    // Remove shadow effects for all types
    XmlHelper.removeByTagName(element, 'a:outerShdw');
    XmlHelper.removeByTagName(element, 'a:innerShdw');

    // Remove 3D effects for all types
    XmlHelper.removeByTagName(element, 'a:scene3d');
    XmlHelper.removeByTagName(element, 'a:sp3d');

    // Remove reflection effects
    XmlHelper.removeByTagName(element, 'a:reflection');

    // Remove glow effects
    XmlHelper.removeByTagName(element, 'a:glow');

    // Apply type-specific cleanup
    switch (elementType) {
      case 'picture':
      case 'imageFilledShape':
        // Remove artistic effects
        XmlHelper.removeByTagName(element, 'a:duotone');
        XmlHelper.removeByTagName(element, 'a:artistic');
        XmlHelper.removeByTagName(element, 'a:colorChange');
        XmlHelper.removeByTagName(element, 'a:softEdge');
        XmlHelper.removeByTagName(element, 'a:sketch');
        XmlHelper.removeByTagName(element, 'ask:lineSketchStyleProps');
        XmlHelper.removeByTagName(element, 'p:style');
        break;

      case 'icon':
      case 'smartArt':
        // For icons and SmartArt, preserve basic structure but remove decorative effects
        XmlHelper.removeByTagName(element, 'a:prstTxWarp');
        XmlHelper.removeByTagName(element, 'a:sketch');
        XmlHelper.removeByTagName(element, 'ask:lineSketchStyleProps');
        break;

      case 'chart':
        // For charts, be more conservative - only remove extreme effects
        XmlHelper.removeByTagName(element, 'c:view3D');
        XmlHelper.removeByTagName(element, 'c:perspective');
        XmlHelper.removeByTagName(element, 'a:prstTxWarp');
        break;

      case 'vectorShape':
        // For vector shapes/drawings, preserve most styling but remove extreme effects
        XmlHelper.removeByTagName(element, 'a:prstTxWarp');
        XmlHelper.removeByTagName(element, 'ask:lineSketchStyleProps');
        break;

      case '3dModel':
        // For 3D models, preserve 3D-specific elements but remove unnecessary styling
        XmlHelper.removeByTagName(element, 'a:prstTxWarp');
        break;

      default:
        // For unknown types, be conservative - only remove clearly problematic effects
        XmlHelper.removeByTagName(element, 'a:prstTxWarp');
        XmlHelper.removeByTagName(element, 'a:sketch');
        XmlHelper.removeByTagName(element, 'ask:lineSketchStyleProps');
    }
  }

  /**
   * Removes text formatting effects and converts underlines to bold
   *
   * @param element - The XML element containing text runs to clean up
   */
  static removeTextEffects(element: XmlElement): void {
    // Get all text run property elements
    const textRuns = element.getElementsByTagName('a:rPr');

    // Process each text run
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      // Remove font specification elements
      const latin = textRun.getElementsByTagName('a:latin').item(0);
      const ea = textRun.getElementsByTagName('a:ea').item(0);
      const cs = textRun.getElementsByTagName('a:cs').item(0);

      // Remove all font elements
      XmlHelper.remove(latin);
      XmlHelper.remove(ea);
      XmlHelper.remove(cs);

      // Check for formatting attributes
      const isBold = textRun.getAttribute('b');
      const isUnderlined = textRun.getAttribute('u');

      // Convert underlined text to bold text
      if (textRun && isUnderlined) {
        textRun.removeAttribute('u');
        textRun.setAttribute('b', '1');
      }
    });
  }
}
