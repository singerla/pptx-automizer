import { createHash } from 'crypto';
import IArchive from '../interfaces/iarchive';
import { PptPaths } from './ppt-paths';

/**
 * PowerPoint stores each distinct media file once and lets any number of
 * relations point to the same part. pptx-automizer used to copy an imported
 * image for every shape referencing it, so the same template image on ten
 * slides ended up ten times in ppt/media, bloating the output (see #145).
 *
 * This is a checksum index of the media files an output archive contains.
 * An import can look up its file contents here and re-use an identical file
 * instead of adding another copy. Relations are per shape and unaffected by
 * sharing a target, and media files are never modified in place - a media
 * file is either copied from a template or added by `Automizer.addMedia()`.
 *
 * One instance per root template; media files of source templates are
 * irrelevant, only what ends up in the output is indexed.
 */
export class MediaDeduplicator {
  archive: IArchive;

  /**
   * Checksum of a media file mapped to its filename inside ppt/media.
   */
  private files = new Map<string, string>();

  private indexed: Promise<void>;

  constructor(archive: IArchive) {
    this.archive = archive;
  }

  /**
   * Find a media file with identical contents in the output archive.
   * @param content Contents of the media file about to be copied
   * @returns Filename inside ppt/media, or undefined if it is a new file
   */
  async find(content: Buffer): Promise<string | undefined> {
    await this.indexExistingFiles();
    return this.files.get(MediaDeduplicator.checksum(content));
  }

  /**
   * Add a media file that has been copied into the output archive.
   */
  add(content: Buffer, filename: string): void {
    const checksum = MediaDeduplicator.checksum(content);
    if (!this.files.has(checksum)) {
      this.files.set(checksum, filename);
    }
  }

  /**
   * Index the media files the output archive already contains, e.g. the ones
   * coming with the root template. Runs on first use only.
   */
  private async indexExistingFiles(): Promise<void> {
    if (!this.indexed) {
      this.indexed = this.readExistingFiles();
    }
    return this.indexed;
  }

  private async readExistingFiles(): Promise<void> {
    const mediaFiles = await this.archive.folder(PptPaths.mediaDir);

    for (const mediaFile of mediaFiles) {
      // Directory entries have no filename to relate to
      if (!mediaFile.relativePath || mediaFile.relativePath.endsWith('/')) {
        continue;
      }

      const content = (await this.archive.read(
        mediaFile.name,
        'nodebuffer',
      )) as Buffer;

      this.add(content, mediaFile.relativePath);
    }
  }

  static checksum(content: Buffer): string {
    return createHash('sha1').update(content).digest('hex');
  }
}
