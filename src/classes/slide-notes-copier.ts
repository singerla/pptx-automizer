import { FileHelper } from '../helper/file-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import type HasShapes from './has-shapes';

/**
 * Copies a slide's notesSlide part (if any) and remaps the notesSlide
 * numbering: slideNote numbers differ from slide numbers if the source
 * presentation contains slides without notes.
 *
 * Extracted from HasShapes (ROADMAP Phase 2).
 */
export class SlideNotesCopier {
  constructor(private host: HasShapes) {}

  /**
   * Copy the source slide's notes (if present) to the target slide,
   * remap the mutual relationships and register the content type.
   */
  async copySlideNotes(): Promise<void> {
    const sourceNotesNumber = await this.getSlideNoteSourceNumber();
    if (sourceNotesNumber) {
      await this.copySlideNoteFiles(sourceNotesNumber);
      await this.updateSlideNoteFile(sourceNotesNumber);
      await this.host.contentTypes.appendNotesToContentType(
        this.host.targetNumber,
      );
    }
  }

  /**
   * Find the proper enumeration of the source notesSlide xml file.
   */
  async getSlideNoteSourceNumber(): Promise<number | undefined> {
    const host = this.host;
    const targets = await XmlHelper.getTargetsByRelationshipType(
      host.sourceArchive,
      PptPaths.slideRels(host.sourceNumber),
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    );

    if (targets.length) {
      const targetNumber = targets[0].file
        .replace('../notesSlides/notesSlide', '')
        .replace('.xml', '');
      return Number(targetNumber);
    }
  }

  async copySlideNoteFiles(sourceNotesNumber: number): Promise<void> {
    const host = this.host;
    await FileHelper.zipCopy(
      host.sourceArchive,
      PptPaths.notesSlide(sourceNotesNumber),
      host.targetArchive,
      PptPaths.notesSlide(host.targetNumber),
    );

    await FileHelper.zipCopy(
      host.sourceArchive,
      PptPaths.notesSlideRels(sourceNotesNumber),
      host.targetArchive,
      PptPaths.notesSlideRels(host.targetNumber),
    );
  }

  /**
   * Point the copied notesSlide at the target slide and vice versa.
   */
  async updateSlideNoteFile(sourceNotesNumber: number): Promise<void> {
    const host = this.host;
    await XmlHelper.replaceAttribute(
      host.targetArchive,
      PptPaths.notesSlideRels(host.targetNumber),
      'Relationship',
      'Target',
      PptPaths.relative.slide(host.sourceNumber),
      PptPaths.relative.slide(host.targetNumber),
    );

    await XmlHelper.replaceAttribute(
      host.targetArchive,
      PptPaths.slideRels(host.targetNumber),
      'Relationship',
      'Target',
      PptPaths.relative.notesSlide(sourceNotesNumber),
      PptPaths.relative.notesSlide(host.targetNumber),
    );
  }
}
