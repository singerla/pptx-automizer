import path from 'path';
import fs from 'fs';
import JSZip from 'jszip';
import { vd } from './general-helper';
import { exists, FileHelper, makeDirIfNotExists } from './file-helper';
const extract = require('extract-zip');

class CacheHelperClass {
  dir: string = undefined;
  templatesDir: string;
  outputDir: string;
  isActive: boolean;

  constructor() {}

  setDir(location: string): void {
    this.dir = location + '/';
    this.templatesDir = this.dir + 'templates' + '/';
    this.outputDir = this.dir + 'output' + '/';

    makeDirIfNotExists(this.dir);
    makeDirIfNotExists(this.templatesDir);
    makeDirIfNotExists(this.outputDir);

    this.isActive = true;
  }

  async readFile(file: string) {
    const info = FileHelper.getFileInfo(file);
    const targetDir = this.templatesDir + info.base;

    if (exists(targetDir)) {
      return;
    }

    extract(file, { dir: targetDir }).catch((err) => {
      throw err;
    });
  }
}

export const CacheHelper = new CacheHelperClass();
