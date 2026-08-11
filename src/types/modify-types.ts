import { XmlElement } from './xml-types';

export type ModifyCallback = {
  (element: XmlElement): void;
};
export type ModifyCollectionCallback = {
  (collection: HTMLCollectionOf<XmlElement>): void;
};
/**
 * A Modification is applied to xml elements by ModificationTags.
 * Specify an index if not 0 and put one or more ModifyCallbacks to
 * 'modify' prop.
 */
export type Modification = {
  index?: number;
  last?: boolean;
  all?: boolean;
  collection?: ModifyCollectionCallback;
  children?: ModificationTags;
  modify?: ModifyCallback | ModifyCallback[];
  create?: any;
  isRequired?: boolean;
  fromPrevious?: boolean;
  fromIndex?: number;
  forceCreate?: boolean;
};
/**
 * ModificationTags will specify the target xml tags for your
 * modifications. ModificationTags can be nested by using 'children'.
 */
export type ModificationTags = {
  [tag: string]: Modification;
};
/**
 * Content for a bulleted list: items are stringified; a nested array
 * increases the indentation level of its items.
 */
export type BulletListContent = (string | number | BulletListContent)[];

export type Color = {
  type?: 'schemeClr' | 'srgbClr';
  value: string;
  shade?: string | number; // Shade value (e.g. "50000")
  tint?: string | number;  // Tint value (e.g. "20000")
  alpha?: string | number; // Alpha/transparency value (e.g. "0.5" for 50% opacity)
  // satMod?: string | number; // Saturation modifier
  // lumMod?: string | number; // Luminance modifier
};
export type Border = {
  tag: 'lnL' | 'lnR' | 'lnT' | 'lnB';
  type?: 'solid' | 'sysDot' | string;
  weight?: number;
  color?: Color;
};

/**
 * Outline (a.k.a. line/border) of a shape, picture or placeholder.
 * Mirrors the vocabulary of `Border` (used for table cells), minus the
 * cell-specific `tag`: a shape has exactly one outline.
 */
export type ShapeOutline = {
  /**
   * Line width in EMU, same unit as `Border.weight`.
   * 1pt = 12700 EMU, 1cm = 360000 EMU. Use `PtToEmu()`/`CmToDxa()`.
   */
  weight?: number;
  /** Dash style, rendered as <a:prstDash val="..."/> */
  type?:
    | 'solid'
    | 'dot'
    | 'sysDot'
    | 'dash'
    | 'sysDash'
    | 'lgDash'
    | 'dashDot'
    | 'lgDashDot'
    | 'lgDashDotDot'
    | string;
  color?: Color;
};
export type HyperlinkInfo = {
  target: string | number;
  isInternal?: boolean;
};

export type TextStyle = {
  /** Font size in 1/100 pt (e.g. 1400 = 14pt) */
  size?: number;
  color?: Color;
  isBold?: boolean;
  isItalics?: boolean;
  isUnderlined?: boolean;
  isSuperscript?: boolean;
  isSubscript?: boolean;
  /** Strikethrough, rendered as strike="sngStrike" */
  isStrike?: boolean;
  /** Typeface name, rendered as <a:latin typeface="..."/> */
  fontFamily?: string;
  /** Text highlight (marker pen), rendered as <a:highlight> */
  highlight?: Color;
  hyperlink?: HyperlinkInfo;
};

export type ImageStyle = {
  duotone?: {
    color?: Color;
    prstClr?: string;
    tint?: number;
    satMod?: number;
  };
};
export type ReplaceText = {
  replace: string;
  by: ReplaceTextReplacement | ReplaceTextReplacement[];
};
export type ReplaceTextReplacement = {
  text: string;
  style?: TextStyle;
};
export type ReplaceTextOptions = {
  openingTag: string;
  closingTag: string;
};
