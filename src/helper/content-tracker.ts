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
import { contentTrack } from '../constants/constants';

export class ContentTracker {
  archive: JSZip;
  files: TrackedFiles = {
    'ppt/slideMasters': [],
    'ppt/slideLayouts': [],
    'ppt/slides': [],
    'ppt/charts': [],
    'ppt/embeddings': [],
  };

  relations: TrackedRelations = {
    // '.': [],
    'ppt/slides/_rels': [],
    'ppt/slideMasters/_rels': [],
    'ppt/slideLayouts/_rels': [],
    'ppt/charts/_rels': [],
    'ppt/_rels': [],
    ppt: [],
  };

  relationTags = contentTrack;

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
    await this.trackSlideLayouts();
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
    ).getTrackedRelations('slideMaster');

    await this.addAndAnalyze(slideMasters, 'ppt/slideMasters');
  }

  async trackSlideLayouts() {
    const usedSlideLayouts =
      this.getRelationTag('ppt/slideMasters').getTrackedRelations(
        'slideLayout',
      );

    await this.addAndAnalyze(usedSlideLayouts, 'ppt/slideLayouts');
  }

  async addAndAnalyze(trackedRelations: TrackedRelation[], section: string) {
    const targets = await this.getRelatedContents(trackedRelations);

    targets.forEach((target) => {
      this.trackFile(section + '/' + target.filename);
    });

    const relationTagInfo = this.getRelationTag(section);
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
    relationTagInfo.getTrackedRelations = (role: string) => {
      return relationTagInfo.tags.filter((tag) => tag.role === role);
    };

    for (const relationTag of relationTagInfo.tags) {
      relationTag.targets = relationTag.targets || [];

      if (relationTagInfo.isDir === true) {
        const files = this.files[relationTagInfo.source] || [];
        if (!files.length) {
          // vd('no files');
          // vd(relationTagInfo.source);
        }
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

    relationTag.targets = [...relationTag.targets, ...addTargets];
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

  async collect(
    section: string,
    role: string,
    collection: string[],
  ): Promise<void> {
    const trackedRelations =
      this.getRelationTag(section).getTrackedRelations(role);
    const images = await this.getRelatedContents(trackedRelations);
    images.forEach((image) => collection.push(image.filename));
  }
}

export const contentTracker = new ContentTracker();
