import { ShapeTargetType } from '../types/types';

/**
 * Any numbered part sharing the `ppt/<prefix>s/<prefix><n>.xml` convention.
 */
type NumberedPartPrefix = ShapeTargetType | 'notesSlide';

/**
 * Central builder for all part paths inside a .pptx archive.
 *
 * OOXML paths follow strict conventions (`ppt/slides/slide<n>.xml`,
 * `ppt/slides/_rels/slide<n>.xml.rels`, …). Build them here instead of
 * inlining template strings — one typo-prone convention, one place.
 *
 * Numbering is 1-based and file-name-driven throughout.
 */
export class PptPaths {
  static readonly presentation = 'ppt/presentation.xml';
  static readonly presentationRels = 'ppt/_rels/presentation.xml.rels';
  static readonly contentTypes = '[Content_Types].xml';
  static readonly mediaDir = 'ppt/media';

  /**
   * Generic numbered part: `ppt/<prefix>s/<prefix><n>.xml`.
   * Covers slide, slideMaster, slideLayout and notesSlide.
   */
  static part(prefix: NumberedPartPrefix, n: number): string {
    return `ppt/${prefix}s/${prefix}${n}.xml`;
  }

  /**
   * Relationships file of a numbered part:
   * `ppt/<prefix>s/_rels/<prefix><n>.xml.rels`.
   */
  static partRels(prefix: NumberedPartPrefix, n: number): string {
    return `ppt/${prefix}s/_rels/${prefix}${n}.xml.rels`;
  }

  static slide(n: number): string {
    return PptPaths.part('slide', n);
  }

  static slideRels(n: number): string {
    return PptPaths.partRels('slide', n);
  }

  static slideMaster(n: number): string {
    return PptPaths.part('slideMaster', n);
  }

  static slideMasterRels(n: number): string {
    return PptPaths.partRels('slideMaster', n);
  }

  static slideLayout(n: number): string {
    return PptPaths.part('slideLayout', n);
  }

  static slideLayoutRels(n: number): string {
    return PptPaths.partRels('slideLayout', n);
  }

  static notesSlide(n: number): string {
    return PptPaths.part('notesSlide', n);
  }

  static notesSlideRels(n: number): string {
    return PptPaths.partRels('notesSlide', n);
  }

  static theme(n: number | string): string {
    return `ppt/theme/theme${n}.xml`;
  }

  /**
   * Numbered part in `ppt/charts/`: pass 'chart', 'chartEx', 'style'
   * or 'colors' as name.
   */
  static chartPart(name: string, n: number): string {
    return `ppt/charts/${name}${n}.xml`;
  }

  static chartPartRels(name: string, n: number): string {
    return `ppt/charts/_rels/${name}${n}.xml.rels`;
  }

  static media(filename: string): string {
    return `ppt/media/${filename}`;
  }

  static embedding(filename: string): string {
    return `ppt/embeddings/${filename}`;
  }

  /**
   * `[Content_Types].xml` PartName attributes require a leading slash.
   */
  static partName(path: string): string {
    return `/${path}`;
  }

  /**
   * Relative targets as used inside `_rels` files
   * (relative to the referencing part's directory).
   */
  static relative = {
    slide: (n: number): string => `../slides/slide${n}.xml`,
    notesSlide: (n: number): string => `../notesSlides/notesSlide${n}.xml`,
    slideLayout: (n: number): string => `../slideLayouts/slideLayout${n}.xml`,
    slideMaster: (n: number): string => `../slideMasters/slideMaster${n}.xml`,
    theme: (n: number | string): string => `../theme/theme${n}.xml`,
  };
}
