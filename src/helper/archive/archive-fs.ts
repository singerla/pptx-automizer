import { ArchiveError } from '../../errors';
import Archive from './archive';
import { promises as fsPromises } from 'fs';
import JSZip, { InputType } from 'jszip';
import {
  ArchiveParams,
  AutomizerParams,
} from '../../types/types';
import IArchive, { ArchivedFile } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import ArchiveJszip from './archive-jszip';
import {
  copyDir,
  ensureDirectoryExistence,
  exists,
  FileHelper,
  makeDirIfNotExists,
} from '../file-helper';
import extract from 'extract-zip';
import { compressFolder } from '../jszip-helper';

export default class ArchiveFs extends Archive implements IArchive {
  archive: boolean;
  params: ArchiveParams;
  dir: string = undefined;
  templatesDir: string;
  templateDir: string;
  outputDir: string;
  workDir: string;
  isActive: boolean;
  isRoot: boolean;
  filename: string;
  private opened: Promise<this>;

  constructor(filename: string, params: ArchiveParams) {
    super(filename, params);
  }

  /**
   * Extracts the template and prepares the work dir on first use.
   * Idempotent; every method touching the extracted files awaits this.
   */
  private ensureOpen(): Promise<this> {
    if (this.archive) {
      return Promise.resolve(this);
    }
    this.opened = this.opened ?? this.initialize();
    return this.opened;
  }

  private async initialize() {
    this.setPaths();

    await this.assertDirs();
    await this.extractFile(this.filename);

    if (!this.params.name) {
      await this.prepareWorkDir(this.filename);
      this.isRoot = true;
    }

    this.archive = true;

    return this;
  }

  setPaths(): void {
    this.dir = this.params.baseDir + '/';
    this.templatesDir = this.dir + 'templates' + '/';
    this.outputDir = this.dir + 'output' + '/';
    this.templateDir = undefined;
    this.workDir = this.outputDir + this.params.workDir + '/';
  }

  async assertDirs(): Promise<void> {
    makeDirIfNotExists(this.dir);
    makeDirIfNotExists(this.templatesDir);
    makeDirIfNotExists(this.outputDir);
    makeDirIfNotExists(this.workDir);
  }

  async extractFile(file: string) {
    const targetDir = this.getTemplateDir(file);

    if (exists(targetDir)) {
      return;
    }

    await extract(file, { dir: targetDir }).catch((err) => {
      throw err;
    });
  }

  getTemplateDir(file: string): string {
    const info = FileHelper.getFileInfo(file);
    this.templateDir = this.templatesDir + info.base + '/';
    return this.templateDir;
  }

  async prepareWorkDir(templateDir: string) {
    await this.cleanupWorkDir();

    const fromTemplate = this.getTemplateDir(templateDir);
    await copyDir(fromTemplate, this.workDir);
  }

  fileExists(file: string) {
    return exists(this.getPath(file));
  }

  async folder(dir: string): Promise<ArchivedFile[]> {
    await this.ensureOpen();
    const path = this.getPath(dir);
    const files: ArchivedFile[] = [];

    if (!exists(path)) {
      return files;
    }

    const entries = await fsPromises.readdir(path, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isDirectory()) {
        files.push({
          name: dir + '/' + entry.name,
          relativePath: entry.name,
        });
      }
    }

    return files;
  }

  async read(file: string): Promise<string | Buffer> {
    await this.ensureOpen();

    const path = this.getPath(file);
    return await fsPromises.readFile(path);
  }

  getPath(file: string): string {
    if (this.isRoot) {
      return this.workDir + file;
    }
    return this.templateDir + file;
  }

  async write(file: string, data: string | Buffer): Promise<this> {
    await this.ensureOpen();
    const filename = this.workDir + file;
    ensureDirectoryExistence(filename);
    await fsPromises.writeFile(filename, data);
    return this;
  }

  async remove(file: string): Promise<void> {
    await this.ensureOpen();
    const path = this.getPath(file);
    if (exists(path)) {
      await fsPromises.unlink(path);
    }
  }

  async output(location: string, params: AutomizerParams): Promise<void> {
    await this.ensureOpen();
    await this.writeBuffer(this);
    this.setOptions(params);

    if (exists(location)) {
      await fsPromises.rm(location);
    }

    await compressFolder(this.workDir, location, this.options);

    if (this.params.cleanupWorkDir === true) {
      await this.cleanupWorkDir();
    }
  }

  async cleanupWorkDir(): Promise<void> {
    if (!exists(this.workDir)) {
      return;
    }

    await fsPromises.rm(this.workDir, { recursive: true, force: true });
  }

  async readXml(file: string): Promise<XmlDocument> {
    const isBuffered = this.fromBuffer(file);

    if (!isBuffered) {
      const buffer = await this.read(file);
      if (!buffer) {
        throw new ArchiveError('No buffer for file: ' + file, { file });
      }

      const xmlString = buffer.toString();
      const XmlDocument = this.parseXml(xmlString);
      this.toBuffer(file, XmlDocument);

      return XmlDocument;
    } else {
      return isBuffered.content;
    }
  }

  writeXml(file: string, XmlDocument: XmlDocument): void {
    this.toBuffer(file, XmlDocument);
  }

  /**
   * Used for worksheets only
   **/
  async extract(file: string): Promise<ArchiveJszip> {
    const contents = (await this.read(file)) as Buffer;
    const zip = new JSZip();
    const newArchive = new ArchiveJszip(file, this.params);
    newArchive.archive = await zip.loadAsync(contents as unknown as InputType);
    return newArchive;
  }
}
