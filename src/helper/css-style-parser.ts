/**
 * Parsing of the inline CSS subset that WYSIWYG editors (CKEditor, TinyMCE)
 * emit inside `style="..."` attributes, mapped onto `TextStyle`.
 *
 * Deliberately standalone (no XML/DOM imports) so it can be reused by both the
 * HTML converter and the color helpers without an import cycle.
 *
 * Units: OOXML wants font sizes in 1/100 pt and colors as 6-digit RRGGBB hex
 * without a leading `#`. CSS gives us px, pt, named colors and `rgb()`.
 */

/** 1px at the CSS reference resolution of 96dpi is 0.75pt. */
const PX_TO_PT = 0.75;

/**
 * Named CSS colors. Not the full 148-name set — the basic 16 plus the names
 * that actually turn up in editor palettes and hand-written HTML.
 */
const NAMED_COLORS: Record<string, string> = {
  aqua: '00FFFF',
  black: '000000',
  blue: '0000FF',
  brown: 'A52A2A',
  cyan: '00FFFF',
  darkblue: '00008B',
  darkgray: 'A9A9A9',
  darkgreen: '006400',
  darkgrey: 'A9A9A9',
  darkred: '8B0000',
  fuchsia: 'FF00FF',
  gold: 'FFD700',
  gray: '808080',
  green: '008000',
  grey: '808080',
  indigo: '4B0082',
  lightblue: 'ADD8E6',
  lightgray: 'D3D3D3',
  lightgreen: '90EE90',
  lightgrey: 'D3D3D3',
  lime: '00FF00',
  magenta: 'FF00FF',
  maroon: '800000',
  navy: '000080',
  olive: '808000',
  orange: 'FFA500',
  pink: 'FFC0CB',
  purple: '800080',
  red: 'FF0000',
  silver: 'C0C0C0',
  teal: '008080',
  turquoise: '40E0D0',
  violet: 'EE82EE',
  white: 'FFFFFF',
  yellow: 'FFFF00',
};

/**
 * Split an inline `style` attribute into a lowercase property → value map.
 * Tolerates missing trailing semicolons and stray whitespace.
 */
export const parseInlineCss = (styleAttr: string): Record<string, string> => {
  const declarations: Record<string, string> = {};

  if (!styleAttr) {
    return declarations;
  }

  styleAttr.split(';').forEach((declaration) => {
    const separator = declaration.indexOf(':');
    if (separator === -1) {
      return;
    }
    const property = declaration.slice(0, separator).trim().toLowerCase();
    const value = declaration.slice(separator + 1).trim();
    if (property && value) {
      declarations[property] = value;
    }
  });

  return declarations;
};

/**
 * Normalize any CSS color notation to 6-digit RRGGBB hex without `#`.
 * Returns undefined for notations OOXML cannot express as a plain srgbClr
 * (gradients, `currentColor`, `transparent`, …) so callers can skip them
 * instead of writing an invalid value.
 */
export const normalizeCssColor = (input: string): string | undefined => {
  if (!input) {
    return undefined;
  }

  const value = input.trim().toLowerCase();

  if (NAMED_COLORS[value]) {
    return NAMED_COLORS[value];
  }

  const hexMatch = value.match(/^#?([0-9a-f]{3,8})$/);
  if (hexMatch) {
    const hex = hexMatch[1];
    if (hex.length === 3 || hex.length === 4) {
      // #abc → aabbcc (a 4th digit is alpha, which srgbClr carries separately)
      return hex
        .slice(0, 3)
        .split('')
        .map((char) => char + char)
        .join('')
        .toUpperCase();
    }
    if (hex.length === 6 || hex.length === 8) {
      return hex.slice(0, 6).toUpperCase();
    }
    return undefined;
  }

  // rgb(1,2,3) / rgba(1,2,3,.5) / rgb(1 2 3 / 50%) — alpha is dropped
  const rgbMatch = value.match(/^rgba?\(([^)]+)\)$/);
  if (rgbMatch) {
    const parts = rgbMatch[1]
      .split(/[,/\s]+/)
      .map((part) => part.trim())
      .filter((part) => part !== '');

    if (parts.length < 3) {
      return undefined;
    }

    const channels = parts.slice(0, 3).map((part) => {
      const numeric = parseFloat(part);
      if (isNaN(numeric)) {
        return NaN;
      }
      // Percentages are relative to 255
      const absolute = part.endsWith('%') ? (numeric / 100) * 255 : numeric;
      return Math.min(255, Math.max(0, Math.round(absolute)));
    });

    if (channels.some((channel) => isNaN(channel))) {
      return undefined;
    }

    return channels
      .map((channel) => channel.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }

  return undefined;
};

/**
 * Convert a CSS font-size to OOXML's 1/100 pt.
 *
 * Supports `px` (96dpi), `pt`, and unitless numbers (treated as px, which is
 * what browsers do for legacy `size` attributes). Relative units (`em`, `rem`,
 * `%`, keywords like `larger`) are intentionally *not* resolved: without the
 * cascade there is no base to resolve them against, and guessing would silently
 * resize text. They return undefined, leaving the size inherited.
 */
export const parseCssFontSize = (input: string): number | undefined => {
  if (!input) {
    return undefined;
  }

  const match = input.trim().toLowerCase().match(/^(-?[\d.]+)\s*(px|pt)?$/);
  if (!match) {
    return undefined;
  }

  const numeric = parseFloat(match[1]);
  if (isNaN(numeric) || numeric <= 0) {
    return undefined;
  }

  const points = match[2] === 'pt' ? numeric : numeric * PX_TO_PT;

  return Math.round(points * 100);
};

/** CSS `font-weight` → bold. Numeric weights are bold from 600 up. */
export const parseCssFontWeight = (input: string): boolean | undefined => {
  const value = input.trim().toLowerCase();

  if (value === 'bold' || value === 'bolder') {
    return true;
  }
  if (value === 'normal' || value === 'lighter') {
    return false;
  }

  const numeric = parseInt(value, 10);
  if (isNaN(numeric)) {
    return undefined;
  }

  return numeric >= 600;
};

/** CSS `font-style` → italics. */
export const parseCssFontStyle = (input: string): boolean | undefined => {
  const value = input.trim().toLowerCase();

  if (value === 'italic' || value === 'oblique') {
    return true;
  }
  if (value === 'normal') {
    return false;
  }

  return undefined;
};

/**
 * CSS `text-decoration` (or `text-decoration-line`) → underline / strike.
 * The shorthand can carry color and style too, so this scans for keywords
 * rather than matching the whole value.
 */
export const parseCssTextDecoration = (
  input: string,
): { isUnderlined?: boolean; isStrike?: boolean } => {
  const value = input.trim().toLowerCase();
  const result: { isUnderlined?: boolean; isStrike?: boolean } = {};

  if (value.includes('none')) {
    return { isUnderlined: false, isStrike: false };
  }
  if (value.includes('underline')) {
    result.isUnderlined = true;
  }
  if (value.includes('line-through')) {
    result.isStrike = true;
  }

  return result;
};

/**
 * CSS `font-family` → a single OOXML typeface name. Takes the first family of
 * the stack (OOXML has no fallback list) and strips quotes.
 */
export const parseCssFontFamily = (input: string): string | undefined => {
  const first = input.split(',')[0]?.trim();
  if (!first) {
    return undefined;
  }

  const unquoted = first.replace(/^["']|["']$/g, '').trim();

  return unquoted === '' ? undefined : unquoted;
};

/** CSS `text-align` → OOXML `algn` value. */
export const parseCssTextAlign = (
  input: string,
): 'l' | 'ctr' | 'r' | 'just' | undefined => {
  switch (input.trim().toLowerCase()) {
    case 'left':
    case 'start':
      return 'l';
    case 'center':
      return 'ctr';
    case 'right':
    case 'end':
      return 'r';
    case 'justify':
      return 'just';
    default:
      return undefined;
  }
};
