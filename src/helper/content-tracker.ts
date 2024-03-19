import {
  Target,
  TrackedFiles,
  TrackedRelation,
  TrackedRelationInfo,
  TrackedRelations,
  TrackedRelationTag,
} from '../types/types';
import { FileHelper } from './file-helper';
import { XmlHelper } from './xml-helper';
import { RelationshipAttribute } from '../types/xml-types';
import { contentTrack } from '../constants/constants';
import IArchive from '../interfaces/iarchive';
import { vd } from './general-helper';

export class ContentTracker {
  archive: IArchive;
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

  relationTags = contentTrack();

  constructor() {}

  reset(): void {
    ['files', 'relations'].forEach((section) =>
      Object.keys(this[section]).forEach((key) => {
        this[section][key] = [];
      }),
    );
    this.relationTags = contentTrack();
  }

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

  async analyzeContents(archive: IArchive) {
    this.setArchive(archive);

    await this.analyzeRelationships();
    await this.trackSlideMasters();
    await this.trackSlideLayouts();
  }

  setArchive(archive: IArchive) {
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
      (element, rels) => {
        rels.push({
          file,
          filename,
          rId: element.getAttribute(attribute),
          type: relationTag.type,
        });
      },
      relationTag.tag,
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
      addTarget.getRelatedContent = this.addRelatedContent(
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

  addRelatedContent(relationTagInfo: TrackedRelationTag, addTarget: Target) {
    return async () => {
      if (addTarget.relatedContent) return addTarget.relatedContent;

      const relationsFile =
        relationTagInfo.isDir === true
          ? relationTagInfo.relationsKey + '/' + addTarget.filename + '.rels'
          : relationTagInfo.relationsKey;

      const relationTarget = await XmlHelper.getRelationshipItems(
        this.archive,
        relationsFile,
        (element, rels) => {
          const rId = element.getAttribute('Id');

          if (rId === addTarget.rId) {
            const target = element.getAttribute('Target');
            const targetMode = element.getAttribute('TargetMode');
            const fileInfo = FileHelper.getFileInfo(target);

            if (targetMode !== 'External') {
              rels.push({
                file: target,
                filename: fileInfo.base,
                rId: rId,
                type: element.getAttribute('Type'),
              });
            }
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
    collection?: string[],
  ): Promise<string[]> {
    collection = collection || [];
    const trackedRelationTag = this.getRelationTag(section);
    const trackedRelations = trackedRelationTag.getTrackedRelations(role);

    const relatedTargets = await this.getRelatedContents(trackedRelations);
    relatedTargets.forEach((relatedTarget) =>
      collection.push(relatedTarget.filename),
    );

    return collection;
  }

  filterRelations(section: string, target: string): TrackedRelationInfo[] {
    const relations = this.relations[section];
    return relations.filter((rel) => rel.attributes.Target === target);
  }
}

export const contentTracker = new ContentTracker();
