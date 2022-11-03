import path from 'path';
import fs from 'fs';
import JSZip from 'jszip';
import { vd } from './general-helper';
const extract = require('extract-zip');

export default class CacheHelper {
  dir: string;
  currentLocation: string;
  currentSubDir: string;

  constructor(dir: string) {
    this.dir = dir;
  }

  setLocation(location: string) {
    this.currentLocation = location;
    const baseName = path.basename(this.currentLocation);
    this.currentSubDir = this.dir + '/' + baseName;
    return this;
  }

  store() {
    extract(this.currentLocation, { dir: this.currentSubDir }).catch((err) => {
      throw err;
    });
    return this;
  }
}
