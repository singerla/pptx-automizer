import { TextStyle } from '../types/modify-types';

/** Automatic numbering scheme for ordered lists, rendered as <a:buAutoNum>. */
export type MultiTextAutoNumberType =
  | 'arabicPeriod'
  | 'arabicParenR'
  | 'alphaLcPeriod'
  | 'alphaUcPeriod'
  | 'romanLcPeriod'
  | 'romanUcPeriod'
  | string;

export interface MultiTextRun {
  text?: string;
  style?: TextStyle;
  /**
   * Renders a soft line break (`<a:br/>`) instead of a text run. PPTX has no
   * in-run line break: the break is an element between two runs. A run may
   * either break or carry text, not both.
   */
  break?: boolean;
}

export interface MultiTextParagraph {
  paragraph: {
    /**
     * Indentation/bullet level, 0-based, 0-8 — mirrors the OOXML `lvl`
     * attribute. Level 0 is the first (outermost) level; whether it shows a
     * bullet is decided by `bullet`, not by the level.
     */
    level?: number;
    bullet?: boolean; // Whether to show a bullet or not
    /**
     * Bullet flavour: a literal character (default), automatic numbering for
     * ordered lists, or explicitly none. Ignored when `bullet` is false.
     */
    bulletType?: 'char' | 'number';
    /** Bullet glyph for `bulletType: 'char'` (default '•') */
    bulletChar?: string;
    /** Numbering scheme for `bulletType: 'number'` (default 'arabicPeriod') */
    autoNumberType?: MultiTextAutoNumberType;
    alignment?: 'l' | 'ctr' | 'r' | 'just'; // Text alignment
    lineSpacing?: number; // Line spacing in points
    spaceBefore?: number; // Space before paragraph in points
    spaceAfter?: number; // Space after paragraph in points
    indent?: number; // Custom indentation in EMU (if different from level)
    marginLeft?: number; // Left margin in EMU
  };
  text?: string;
  style?: TextStyle;
  textRuns?: MultiTextRun[];
}
