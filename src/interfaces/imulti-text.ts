import { TextStyle } from '../types/modify-types';

export interface MultiTextParagraph {
  paragraph: {
    level?: number; // Indentation/bullet level (0 = no bullet, 1+ = bullet levels)
    bullet?: boolean; // Whether to show a bullet or not
    alignment?: 'left' | 'center' | 'right' | 'justify'; // Text alignment
    lineSpacing?: number; // Line spacing in points
    spaceBefore?: number; // Space before paragraph in points
    spaceAfter?: number; // Space after paragraph in points
    indent?: number; // Custom indentation in points (if different from level)
    marginLeft?: number; // Left margin in points
  }
  text?: string;
  style?: TextStyle;
  textRuns: {
    text: string;
    style?: TextStyle;
  }[];
}
