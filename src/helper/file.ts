import fs from 'fs'
import JSZip from 'jszip'

export default class FileHelper {

  static readFile(location:string): Promise<Buffer> {
    return fs.promises.readFile(location)
  }

  static extractFromArchive(archive: JSZip, file: string): Promise<string> {
    if(archive.files[file] === undefined) {
      throw new Error('Archived file not found: ' + file)
    }
    return archive.files[file].async('string')
  }


  static extractFileContent(file: any): Promise<JSZip>{
    const zip = new JSZip();
    return zip.loadAsync(file)
  }

	/**
	 * Copies a file from one archive to another. The new file can have a different name to the origin.
	 * @param {string} sourceArchive - Source archive
	 * @param {string} sourceFile - file path and name inside source archive
   * @param {string} targetArchive - Target archive
	 * @param {string} targetFile - file path and name inside target archive
	 * @return {JSZip} targetArchive as an instance of JSZip
	 */
  static async zipCopy(sourceArchive: JSZip, sourceFile:string, targetArchive: JSZip, targetFile?:string): Promise<JSZip> {
    if(sourceArchive.files[sourceFile] === undefined) {
      throw new Error('File not found: ' + sourceFile)
    }

    let content = sourceArchive.files[sourceFile].async('nodebuffer')
    return targetArchive.file(targetFile || sourceFile, content)
  }

  static writeOutputFile(location: string, content: Buffer): void {
    fs.writeFile(location, content, function(err: { message: any }) {
      if(err) {
        throw new Error(`Error writing output: ${err.message}`)
      } else {
        console.log(`output: ${location}`)
      }
    })
  }
}