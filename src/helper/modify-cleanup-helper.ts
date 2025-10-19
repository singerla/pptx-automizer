import { XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { XmlSlideHelper } from './xml-slide-helper';
import { Color, ModifyColorHelper } from '../index';

/**
 * Helper class for cleaning and clearing PowerPoint effects in XML structure
 */
export default class ModifyCleanupHelper {
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
  static removeEffects(element: XmlElement): void {
    this.removeShapeStyle(element);
    this.removeColorAdjustment(element);
    this.removeShadowEffects(element);
    this.remove3dEffects(element);
    this.removeFillEffects(element);
    this.removeTextEffects(element);

    // Determine the visual type of the element (picture, chart, etc.)
    // Apply type-specific cleanup
    const elementType = XmlSlideHelper.getElementVisualType(element);

    switch (elementType) {
      case 'picture':
      case 'svgImage':
      case 'pictogram':
        // Remove artistic effects
        [
          'a:duotone',
          'a:artistic',
          'a:colorChange',
          'a:softEdge',
          'a:prstTxWarp',
          'a:sketch',
          'ask:lineSketchStyleProps',
        ].forEach((tag) => {
          XmlHelper.removeByTagName(element, tag);
        });
        break;

      case 'chart':
        // For charts, be more conservative - only remove extreme effects
        ['c:view3D', 'c:perspective', 'a:prstTxWarp'].forEach((tag) => {
          XmlHelper.removeByTagName(element, tag);
        });
        break;

      default:
        // For unknown types, be conservative - only remove clearly problematic effects
        ['a:prstTxWarp', 'a:sketch', 'ask:lineSketchStyleProps'].forEach(
          (tag) => {
            XmlHelper.removeByTagName(element, tag);
          },
        );
    }
  }

  /**
   * Removes text formatting effects
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
    });
  }

  static clearTextUnderlineToBold(element: XmlElement): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      const isUnderlined = textRun.getAttribute('u');
      // Convert underlined text to bold text
      if (textRun && isUnderlined) {
        textRun.removeAttribute('u');
        textRun.setAttribute('b', '1');
      }
    });
  }

  static clearTextItalicsToBold(element: XmlElement): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      const isItalics = textRun.getAttribute('i');
      // Convert italics text to bold text
      if (textRun && isItalics) {
        textRun.removeAttribute('i');
        textRun.setAttribute('b', '1');
      }
    });
  }

  static clearTextUnderline(element: XmlElement): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      textRun?.removeAttribute('u');
    });
  }

  static clearTextBold(element: XmlElement): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      textRun?.removeAttribute('b');
    });
  }

  static clearTextSize(element: XmlElement): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      textRun?.removeAttribute('sz');
    });
  }

  static clearTextColor(element: XmlElement, color?: Color): void {
    const textRuns = element.getElementsByTagName('a:rPr');
    XmlHelper.modifyCollection(textRuns, (textRun: XmlElement) => {
      if(color) {
        ModifyColorHelper.solidFill(color, 0)(textRun);
      } else {
        // Remove all color-related elements from text run properties
        const solidFill = textRun.getElementsByTagName('a:solidFill')[0];
        if (solidFill) {
          XmlHelper.remove(solidFill);
        }

        const gradFill = textRun.getElementsByTagName('a:gradFill')[0];
        if (gradFill) {
          XmlHelper.remove(gradFill);
        }

        const pattFill = textRun.getElementsByTagName('a:pattFill')[0];
        if (pattFill) {
          XmlHelper.remove(pattFill);
        }

        const noFill = textRun.getElementsByTagName('a:noFill')[0];
        if (noFill) {
          XmlHelper.remove(noFill);
        }

        // Remove highlight color
        const highlight = textRun.getElementsByTagName('a:highlight')[0];
        if (highlight) {
          XmlHelper.remove(highlight);
        }
      }
    });
  }

  static removeFillEffects(element: XmlElement) {
    // Remove reflection, glow effects, gradients and complex fill effects
    ['a:reflection', 'a:glow', 'a:gradFill', 'a:pattFill', 'a:grpFill'].forEach(
      (tag) => {
        XmlHelper.removeByTagName(element, tag);
      },
    );

    ['a:scene3d', 'a:sp3d'].forEach((tag) => {
      XmlHelper.removeByTagName(element, tag);
    });
  }

  static remove3dEffects(element: XmlElement) {
    ['a:scene3d', 'a:sp3d'].forEach((tag) => {
      XmlHelper.removeByTagName(element, tag);
    });
  }

  static removeShadowEffects(element: XmlElement) {
    // Remove shadow effects for all types
    [
      'a:outerShdw',
      'a:innerShdw',
      'a:gradFill',
      'a:pattFill',
      'a:grpFill',
    ].forEach((tag) => {
      XmlHelper.removeByTagName(element, tag);
    });
  }

  static removeColorAdjustment(element: XmlElement) {
    // Remove color adjustments and transformations
    [
      'a:alpha',
      'a:alphaInv',
      'a:alphaMod',
      'a:alphaOff',
      'a:blue',
      'a:blueMod',
      'a:blueOff',
      'a:comp',
      'a:gamma',
      'a:gray',
      'a:green',
      'a:greenMod',
      'a:greenOff',
      'a:hue',
      'a:hueMod',
      'a:hueOff',
      'a:inv',
      'a:invGamma',
      'a:lum',
      'a:lumMod',
      'a:lumOff',
      'a:red',
      'a:redMod',
      'a:redOff',
      'a:sat',
      'a:satMod',
      'a:satOff',
      'a:shade',
      'a:tint',
    ].forEach((tag) => {
      XmlHelper.removeByTagName(element, tag);
    });
  }
}
