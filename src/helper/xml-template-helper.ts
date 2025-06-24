import JSZip from 'jszip';
import { Target } from '../types/types';
import {
  ElementInfo,
  SlideInfo,
  TemplateSlideInfo,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import IArchive from '../interfaces/iarchive';
import { XmlRelationshipHelper } from './xml-relationship-helper';
import { XmlSlideHelper } from './xml-slide-helper';
import { vd } from './general-helper';

export class XmlTemplateHelper {
  archive: IArchive;
  relType: string;
  relTypeNotes: string;
  path: string;
  defaultSlideName: string;

  constructor(archive: IArchive) {
    this.relType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
    this.relTypeNotes =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
    this.archive = archive;
    this.path = 'ppt/_rels/presentation.xml.rels';
    this.defaultSlideName = 'untitled';
  }

  async getCreationIds(): Promise<SlideInfo[]> {
    const archive = this.archive;
    const relationships = await XmlHelper.getTargetsByRelationshipType(
      archive,
      this.path,
      this.relType,
    );

    // ToDo: The slide list is based on the relations from this.path
    // which contains non-visible slides, too.
    // Should be either:
    //  a.) remove unused slides on generation
    //  b.) use slides list from 'p:sldIdLst' in `ppt/presentation.xml`
    vd(this.relType)
    vd(this.path)

    const creationIds: SlideInfo[] = [];
    for (const slideRel of relationships) {
      try {
        const slideXml = await XmlHelper.getXmlFromArchive(
          archive,
          'ppt/' + slideRel.file,
        );
        if (!slideXml) {
          console.warn(`slideXml is undefined for file ${slideRel.file}`);
          continue;
        }

        const slideHelper = new XmlSlideHelper(slideXml);
        const creationIdSlide = slideHelper.getSlideCreationId();
        if (!creationIdSlide) {
          console.warn(`No creationId found in ${slideRel.file}`);
        }

        const slideInfo = await this.getSlideInfo(
          slideXml,
          archive,
          slideRel.file,
        );

        creationIds.push({
          id: creationIdSlide,
          number: this.parseSlideRelFile(slideRel.file),
          elements: slideHelper.getAllElements(),
          info: slideInfo,
        });
      } catch (err) {
        console.error(
          `An error occurred while processing ${slideRel.file}:`,
          err,
        );
      }
    }

    return creationIds.sort((slideA, slideB) =>
      slideA.number < slideB.number ? -1 : 1,
    );
  }

  parseSlideRelFile(slideRelFile: string): number {
    return Number(slideRelFile.replace('slides/slide', '').replace('.xml', ''));
  }

  async getSlideInfo(
    slideXml: XmlDocument,
    archive,
    slideRelFile: string,
  ): Promise<TemplateSlideInfo> {
    let name;

    const slideNoteRels = await this.getSlideNoteRels(archive, slideRelFile);
    if (slideNoteRels.length > 0) {
      name = await this.getSlideNameFromNotes(archive, slideNoteRels);
    }

    if (!name) {
      name = this.getNameFromSlideInfo(slideXml);
    }

    name = !name ? this.defaultSlideName : name;

    return {
      name: name,
    };
  }

  getNameFromSlideInfo(slideXml: XmlDocument): string {
    const slideTitle = slideXml.getElementsByTagName('p:ph');

    if (slideTitle.length && slideTitle[0].getAttribute('type') === 'title') {
      const titleElement = slideTitle[0].parentNode.parentNode
        .parentNode as XmlElement;
      const nameFragments = this.parseTitleElement(titleElement);

      if (nameFragments.length) {
        return nameFragments.join(' ');
      }
    }
  }

  async getSlideNoteRels(
    archive: IArchive,
    slideRelFile: string,
  ): Promise<Target[]> {
    const relFileName = slideRelFile.replace('slides', '');
    const slideRels = await XmlHelper.getTargetsByRelationshipType(
      archive,
      `ppt/slides/_rels${relFileName}.rels`,
      this.relTypeNotes,
    );
    return slideRels;
  }

  async getSlideNameFromNotes(archive, slideNoteRels): Promise<string> {
    const notesFile = slideNoteRels[0].file.replace('../', '');
    const notesXml = await XmlHelper.getXmlFromArchive(
      archive,
      'ppt/' + notesFile,
    );

    const titleElements = notesXml.getElementsByTagName('a:p');
    if (titleElements.length > 0) {
      const nameFragments = this.parseTitleElement(titleElements[0]);
      if (nameFragments.length) {
        return nameFragments.join('');
      }
    }
  }

  parseTitleElement(titleElement: XmlElement): string[] {
    const nameFragments = [];
    const titleText = titleElement.getElementsByTagName('a:t');

    if (titleText.length) {
      for (const titleTextNode in titleText) {
        if (titleText[titleTextNode].firstChild?.nodeValue) {
          nameFragments.push(titleText[titleTextNode].firstChild.nodeValue);
        }
      }
    }

    return nameFragments;
  }
  /**
   * Returns the slide numbers of a given template as a sorted array of integers.
   * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
   */
  async getAllSlideNumbers(): Promise<number[]> {
    try {
      const archive = this.archive;
      const xmlRelationshipHelper = new XmlRelationshipHelper();
      const allSlides = (await xmlRelationshipHelper.initialize(
        archive,
        'presentation.xml.rels',
        'ppt/_rels',
        'slides/slide',
      )) as Target[];

      // Extract slide numbers from each slide using the 'number' property and sort the array of integers.
      const slideNumbers = allSlides.map((slide) => slide.number);
      slideNumbers.sort((a, b) => a - b);

      return slideNumbers;
    } catch (error) {
      throw new Error(`Error getting slide numbers: ${error.message}`);
    }
  }
}
