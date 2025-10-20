import { Color, ImageStyle } from '../types/modify-types';
import XmlElements from './xml-elements';
import { XmlHelper } from './xml-helper';
import { ShapeBackgroundInfo, XmlElement } from '../types/xml-types';
import { vd } from './general-helper';
import { ModifyShapeHelper } from '../index';

export default class ModifyColorHelper {
  /**
   * Replaces or creates an <a:solidFill> Element.
   * The given elelement must be a <p:spPr> or <a:spPr>
   */
  static solidFill =
    (color: Color, index?: number | 'last') =>
    (element: XmlElement): void => {
      if (!color || !color.type || element?.getElementsByTagName === undefined)
        return;

      ModifyColorHelper.normalizeColorObject(color);

      const solidFills = element.getElementsByTagName('a:solidFill');

      if (!solidFills.length) {
        const solidFill = new XmlElements(element, {
          color: color,
        }).solidFill();

        if (element.firstChild && index && index === 0) {
          element.insertBefore(solidFill, element.firstChild);
        } else {
          element.appendChild(solidFill);
        }
        return;
      }

      const targetIndex = !index
        ? 0
        : index === 'last'
        ? solidFills.length - 1
        : index;

      const solidFill = solidFills[targetIndex] as XmlElement;
      const colorType = new XmlElements(element, {
        color: color,
      }).colorType();

      XmlHelper.sliceCollection(
        solidFill.childNodes as unknown as HTMLCollectionOf<XmlElement>,
        0,
      );
      solidFill.appendChild(colorType);
    };

  static removeNoFill =
    () =>
    (element: XmlElement): void => {
      const hasNoFill = element.getElementsByTagName('a:noFill')[0];
      if (hasNoFill) {
        element.removeChild(hasNoFill);
      }
    };

  static normalizeColorObject = (color: Color) => {
    if (color.value.indexOf('#') === 0) {
      color.value = color.value.replace('#', '');
    }
    if (color.value.toLowerCase().indexOf('rgb(') === 0) {
      // TODO: convert RGB to HEX
      color.value = 'cccccc';
    }
    return color;
  };

  /**
   * Check if the given element has a background which is non-transparent
   * @param element The XML element to check
   * @returns ShapeBackgroundInfo
   */
  static elementHasBackground(element: XmlElement): ShapeBackgroundInfo {
    // Find the shape properties (spPr) element
    const spPr = element.getElementsByTagName('p:spPr')[0] ||
      element.getElementsByTagName('a:spPr')[0];

    if (!spPr) {
      // No shape properties found - assume no background
      return {
        isDark: false
      };
    }

    // Check for different fill types in order of priority

    // 1. Check for solid fill
    const solidFill = spPr.getElementsByTagName('a:solidFill')[0];
    if (solidFill) {
      const schemeClr = solidFill.getElementsByTagName('a:schemeClr')[0];
      const srgbClr = solidFill.getElementsByTagName('a:srgbClr')[0];

      if (schemeClr) {
        const schemeValue = schemeClr.getAttribute('val');
        return {
          color: { type: 'schemeClr', value: schemeValue || 'bg1' },
          isDark: ModifyColorHelper.isSchemeColorDark(schemeValue || 'bg1')
        };
      }

      if (srgbClr) {
        const rgbValue = srgbClr.getAttribute('val');
        return {
          color: { type: 'srgbClr', value: rgbValue || 'FFFFFF' },
          isDark: ModifyColorHelper.isRgbColorDark(rgbValue || 'FFFFFF')
        };
      }

      // Solid fill exists but no color specified - assume light
      return {
        isDark: false
      };
    }

    // 2. Check for gradient fill
    const gradFill = spPr.getElementsByTagName('a:gradFill')[0];
    if (gradFill) {
      // For gradient fills, check the first gradient stop
      const gsLst = gradFill.getElementsByTagName('a:gsLst')[0];
      if (gsLst) {
        const firstGs = gsLst.getElementsByTagName('a:gs')[0];
        if (firstGs) {
          const schemeClr = firstGs.getElementsByTagName('a:schemeClr')[0];
          const srgbClr = firstGs.getElementsByTagName('a:srgbClr')[0];

          if (schemeClr) {
            const schemeValue = schemeClr.getAttribute('val');
            return {
              color: { type: 'schemeClr', value: schemeValue || 'bg1' },
              isDark: ModifyColorHelper.isSchemeColorDark(schemeValue || 'bg1')
            };
          }

          if (srgbClr) {
            const rgbValue = srgbClr.getAttribute('val');
            return {
              color: { type: 'srgbClr', value: rgbValue || 'FFFFFF' },
              isDark: ModifyColorHelper.isRgbColorDark(rgbValue || 'FFFFFF')
            };
          }
        }
      }

      // Gradient fill exists but no color found - assume light
      return {
        isDark: false
      };
    }

    // 3. Check for pattern fill or other fill types
    const pattFill = spPr.getElementsByTagName('a:pattFill')[0];
    if (pattFill) {
      // Pattern fills typically have a foreground color - check that
      const fgClr = pattFill.getElementsByTagName('a:fgClr')[0];
      if (fgClr) {
        const schemeClr = fgClr.getElementsByTagName('a:schemeClr')[0];
        const srgbClr = fgClr.getElementsByTagName('a:srgbClr')[0];

        if (schemeClr) {
          const schemeValue = schemeClr.getAttribute('val');
          return {
            color: { type: 'schemeClr', value: schemeValue || 'bg1' },
            isDark: ModifyColorHelper.isSchemeColorDark(schemeValue || 'bg1')
          };
        }

        if (srgbClr) {
          const rgbValue = srgbClr.getAttribute('val');
          return {
            color: { type: 'srgbClr', value: rgbValue || 'FFFFFF' },
            isDark: ModifyColorHelper.isRgbColorDark(rgbValue || 'FFFFFF')
          };
        }
      }
    }

    // 4. Check for no fill
    const noFill = spPr.getElementsByTagName('a:noFill')[0];
    if (noFill) {
      // No fill - transparent background
      return {
        isDark: false
      };
    }

    // No specific fill found - assume default light background
    return {
      isDark: false
    };
  }

  /**
   * Determines if a scheme color is typically dark and would require light text
   * @param schemeValue The scheme color value (e.g., 'dk1', 'lt1', 'accent1', etc.)
   * @returns true if the color is typically dark
   */
  private static isSchemeColorDark(schemeValue: string): boolean {
    // Classic dark scheme colors that typically require light/white text
    const darkSchemeColors = [
      'dk1',      // Dark 1 (usually black or very dark)
      'dk2',      // Dark 2 (usually dark blue)
      'accent1',  // Often darker colors
      'accent2',
      'accent4',  // Often purple/dark
      'tx1',      // Text 1 (usually dark)
      'tx2'       // Text 2 (usually dark)
    ];

    return darkSchemeColors.includes(schemeValue);
  }

  /**
   * Determines if an RGB color is dark based on luminance
   * @param rgbValue The RGB hex value (e.g., '000000', 'FF0000')
   * @returns true if the color is dark
   */
  private static isRgbColorDark(rgbValue: string): boolean {
    // Remove '#' if present and ensure we have a valid hex string
    const hex = rgbValue.replace('#', '').toUpperCase();

    if (hex.length !== 6) {
      return false; // Invalid hex, assume light
    }

    // Convert hex to RGB
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);

    // Calculate relative luminance using the standard formula
    // Values below 0.5 are considered dark (on a scale of 0-1)
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;

    return luminance < 0.5;
  }
}
