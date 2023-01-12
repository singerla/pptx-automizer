import {
  Target,
  TrackedFiles,
  TrackedRelation,
  TrackedRelationInfo,
  TrackedRelations,
  TrackedRelationTag,
} from '../types/types';
import { Slide } from '../classes/slide';
import { FileHelper } from './file-helper';
import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';
import JSZip from 'jszip';
import { RelationshipAttribute } from '../types/xml-types';

export class ContentTracker {
  archive: JSZip;
  files: TrackedFiles = {
    'ppt/slideMasters': [],
    'ppt/slideMasters/_rels': [],
    'ppt/slides': [],
    'ppt/slides/_rels': [],
    'ppt/charts': [],
    'ppt/charts/_rels': [],
    'ppt/embeddings': [],
  };

  relations: TrackedRelations = {
    // '.': [],
    'ppt/slides/_rels': [],
    'ppt/slideMasters/_rels': [],
    'ppt/charts/_rels': [],
    'ppt/_rels': [],
    ppt: [],
  };

  relationTags: TrackedRelationTag[] = [
    {
      source: 'ppt/presentation.xml',
      relationsKey: 'ppt/_rels/presentation.xml.rels',
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
          tag: 'p:sldMasterId',
          role: 'slideMaster',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
          tag: 'p:sldId',
          role: 'slide',
        },
      ],
    },
    {
      source: 'ppt/slides',
      relationsKey: 'ppt/slides/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
          tag: 'c:chart',
          role: 'chart',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          tag: 'a:blip',
          role: 'image',
          attribute: 'r:embed',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          tag: 'asvg:svgBlip',
          role: 'image',
          attribute: 'r:embed',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          role: 'slideLayout',
          tag: null,
        },
      ],
    },
    {
      source: 'ppt/charts',
      relationsKey: 'ppt/charts/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
          tag: 'c:externalData',
          role: 'externalData',
        },
      ],
    },
    {
      source: 'ppt/slideMasters',
      relationsKey: 'ppt/slideMasters/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          tag: 'p:sldLayoutId',
          role: 'slideLayout',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          tag: 'a:blip',
          role: 'image',
          attribute: 'r:embed',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          tag: 'asvg:svgBlip',
          role: 'image',
          attribute: 'r:embed',
        },
      ],
    },
  ];

  constructor() {}

  trackFile(file: string): void {
    const info = FileHelper.getFileInfo(file);
    if (this.files[info.dir]) {
      this.files[info.dir].push(info.base);
    }
  }

  trackRelation(file: string, attributes: RelationshipAttribute): void {
    const info = FileHelper.getFileInfo(file);
    if (this.relations[info.dir]) {
      this.relations[info.dir].push({
        base: info.base,
        attributes,
      });
    }
  }

  async analyzeContents(archive: JSZip) {
    this.setArchive(archive);

    await this.analyzeRelationships();
    await this.trackSlideMasters();
  }

  setArchive(archive: JSZip) {
    this.archive = archive;
  }

  /**
   * This will be replaced by future slideMaster handling.
   */
  async trackSlideMasters() {
    const slideMasters = this.getRelationTag(
      'ppt/presentation.xml',
    ).getRelationTargets('slideMaster');

    const slideMasterInfo = await this.getRelatedContents(slideMasters);

    slideMasterInfo.forEach((slideMasterInfo) => {
      this.trackFile('ppt/' + slideMasterInfo.file);
      this.trackFile('ppt/_rels/' + slideMasterInfo.file + '.rels');
    });

    const relationTagInfo = this.getRelationTag('ppt/slideMasters');
    await this.analyzeRelationship(relationTagInfo);
  }

  async getRelatedContents(
    trackedRelations: TrackedRelation[],
  ): Promise<Target[]> {
    const relatedContents = [];
    for (const trackedRelation of trackedRelations) {
      for (const target of trackedRelation.targets) {
        const trackedRelationInfo = await target.getRelatedContent();
        relatedContents.push(trackedRelationInfo);
      }
    }
    return relatedContents;
  }

  getRelationTag(source: string): TrackedRelationTag {
    return contentTracker.relationTags.find(
      (relationTag) => relationTag.source === source,
    );
  }

  async analyzeRelationships(): Promise<void> {
    for (const relationTagInfo of this.relationTags) {
      await this.analyzeRelationship(relationTagInfo);
    }
  }

  async analyzeRelationship(
    relationTagInfo: TrackedRelationTag,
  ): Promise<void> {
    relationTagInfo.getRelationTargets = (role: string) => {
      return relationTagInfo.tags.filter((tag) => tag.role === role);
    };

    for (const relationTag of relationTagInfo.tags) {
      if (relationTagInfo.isDir === true) {
        const files = this.files[relationTagInfo.source];
        for (const file of files) {
          await this.pushRelationTagTargets(
            relationTagInfo.source + '/' + file,
            file,
            relationTag,
            relationTagInfo,
          );
        }
      } else {
        const pathInfo = FileHelper.getFileInfo(relationTagInfo.source);
        await this.pushRelationTagTargets(
          relationTagInfo.source,
          pathInfo.base,
          relationTag,
          relationTagInfo,
        );
      }
    }
  }

  async pushRelationTagTargets(
    file: string,
    filename: string,
    relationTag: TrackedRelation,
    relationTagInfo,
  ): Promise<void> {
    const attribute = relationTag.attribute || 'r:id';

    const addTargets = await XmlHelper.getRelationshipItems(
      this.archive,
      file,
      relationTag.tag,
      (element, rels) => {
        rels.push({
          file,
          filename,
          rId: element.getAttribute(attribute),
          type: relationTag.type,
        });
      },
    );

    this.addCreatedRelationsFunctions(
      addTargets,
      contentTracker.relations[relationTagInfo.relationsKey],
      relationTagInfo,
    );

    const existingTargets = relationTag.targets || [];
    relationTag.targets = [...existingTargets, ...addTargets];
  }

  addCreatedRelationsFunctions(
    addTargets: Target[],
    createdRelations: TrackedRelationInfo[],
    relationTagInfo: TrackedRelationTag,
  ): void {
    addTargets.forEach((addTarget) => {
      addTarget.getCreatedContent = this.getCreatedContent(
        createdRelations,
        addTarget,
      );
      addTarget.getRelatedContent = this.getRelatedContent(
        relationTagInfo,
        addTarget,
      );
    });
  }

  getCreatedContent(
    createdRelations: TrackedRelationInfo[],
    addTarget: Target,
  ) {
    return () => {
      return createdRelations.find((relation) => {
        return (
          relation.base === addTarget.filename + '.rels' &&
          relation.attributes?.Id === addTarget.rId
        );
      });
    };
  }

  getRelatedContent(relationTagInfo: TrackedRelationTag, addTarget: Target) {
    return async () => {
      if (addTarget.relatedContent) return addTarget.relatedContent;

      const relationsFile =
        relationTagInfo.isDir === true
          ? relationTagInfo.relationsKey + '/' + addTarget.filename + '.rels'
          : relationTagInfo.relationsKey;

      const relationTarget = await XmlHelper.getRelationshipItems(
        this.archive,
        relationsFile,
        'Relationship',
        (element, rels) => {
          const rId = element.getAttribute('Id');

          if (rId === addTarget.rId) {
            const target = element.getAttribute('Target');
            const fileInfo = FileHelper.getFileInfo(target);

            rels.push({
              file: target,
              filename: fileInfo.base,
              rId: rId,
              type: element.getAttribute('Type'),
            });
          }
        },
      );

      addTarget.relatedContent = relationTarget.find(
        (relationTarget) => relationTarget.rId === addTarget.rId,
      );

      return addTarget.relatedContent;
    };
  }

  dump() {
    console.log(this.files);
    // console.log(this.relations);
  }
}

export const contentTracker = new ContentTracker();
