import { SourceSlideIdentifier } from '../types/types';
import { Slide } from '../classes/slide';
import { FileHelper } from './file-helper';

export type TrackedPresentation = {
  slides: TrackedSlide[];
};

export type TrackedSlide = {
  targetPath: string;
  targetRelsPath: string;
};

export type FileInfo = {
  base: string;
  extension: string;
  dir: string;
};

export class ContentTracker {
  files: Record<string, string[]> = {
    'ppt/slides': [],
    'ppt/slides/_rels': [],
    'ppt/charts': [],
    'ppt/charts/_rels': [],
    'ppt/embeddings': [],
  };

  relations: Record<
    string,
    {
      base: string;
      attribute: string;
      value: string;
    }[]
  > = {
    // '.': [],
    'ppt/slides/_rels': [],
    'ppt/charts/_rels': [],
    'ppt/_rels': [],
    ppt: [],
  };

  constructor() {}

  trackFile(file): void {
    const info = FileHelper.getFileInfo(file);
    if (this.files[info.dir]) {
      this.files[info.dir].push(info.base);
    } else {
      console.log(`Could not track file ${file}`);
    }
  }

  trackRelation(file: string, attribute: string, value: string): void {
    const info = FileHelper.getFileInfo(file);

    if (this.relations[info.dir]) {
      this.relations[info.dir].push({
        base: info.base,
        attribute,
        value,
      });
    } else {
      // console.log(`Could not track relation ${info.dir}`);
    }
  }

  dump() {
    console.log(this.files);
    console.log(this.relations);
  }
}

export const contentTracker = new ContentTracker();
