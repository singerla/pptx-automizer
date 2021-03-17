import fs from 'fs'
import path from 'path'
import JSZip from 'jszip'
import { AutomizerSummary, IPresentationProps } from '../definitions/app'

export default class FileHelper {

  static readFile(location:string): Promise<Buffer> {
    if(!fs.existsSync(location)) {
      throw new Error('File not found: ' + location)
    }
    return fs.promises.readFile(location)
  }

  static extractFromArchive(archive: JSZip, file: string): Promise<string> {
    if(archive === undefined) {
      throw new Error('No files found, expected: ' + file)
    }

    if(archive.files[file] === undefined) {
      throw new Error('Archived file not found: ' + file)
    }
    return archive.files[file].async('string')
  }

  static extractFileContent(file: any): Promise<JSZip>{
    const zip = new JSZip();
    return zip.loadAsync(file)
  }

  static getFileExtension(filename: string): string {
    let extension = path.extname(filename).replace('.', '')
    return extension
  }

	/**
	 * Copies a file from one archive to another. The new file can have a different name to the origin.
	 * @param {JSZip} sourceArchive - Source archive
	 * @param {string} sourceFile - file path and name inside source archive
   * @param {JSZip} targetArchive - Target archive
	 * @param {string} targetFile - file path and name inside target archive
	 * @return {JSZip} targetArchive as an instance of JSZip
	 */
  static async zipCopy(sourceArchive: JSZip, sourceFile:string, targetArchive: JSZip, targetFile?:string, mode?:any): Promise<JSZip> {
    if(sourceArchive.files[sourceFile] === undefined) {
      throw new Error(`Zipped file not found: ${sourceFile}`)
    }

    let content = sourceArchive.files[sourceFile].async('nodebuffer')
    return targetArchive.file(targetFile || sourceFile, content)
  }

  static writeOutputFile(location: string, content: Buffer, automizer: IPresentationProps): AutomizerSummary {
    fs.writeFile(location, content, function(err: { message: any }) {
      if(err) {
        throw new Error(`Error writing output: ${err.message}`)
      }
    })

    let duration: number = (Date.now() - automizer.timer) / 600

    return {
      status: 'finished',
      duration: duration,
      file: location,
      templates: automizer.templates.length,
      slides: automizer.rootTemplate.count('slides'),
      charts: automizer.rootTemplate.count('charts'),
      images: automizer.rootTemplate.count('images'),
    }
  }
}