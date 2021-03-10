import fs from 'fs'
import JSZip, { JSZipFileOptions } from 'jszip'

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

  static extractAllForecefully(filePath: string): void {
    fs.readFile(filePath, function(err, data) {
      if (!err) {
        var path = require('path');
        let dir = filePath + '.unzip'
        fs.rmdirSync(dir, { recursive: true })
        fs.mkdirSync(dir)
        var zip = new JSZip()
        zip.loadAsync(data).then(function(contents) {
          Object.keys(zip.files).forEach(function (filename) {
            let subDir = path.dirname(contents.files[filename].name)
            if(!fs.existsSync(subDir)) {
              fs.mkdirSync(dir + '/' + subDir, { recursive: true })
            }
          })

          Object.keys(zip.files).forEach(function (filename) {
            zip.files[filename].async('string').then(function (fileData) {
              if(contents.files[filename].dir === false) {
                fs.writeFileSync(dir + '/' + filename, fileData)
              }
            })
          })
        })
      }
    })
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