
import fs from 'fs'
import JSZip from 'jszip'

export default class FileHelper  {
  static readFile(location) {
    return fs.promises.readFile(location)
  }

  static extractWorkbook(archive) {
    return archive.files['xl/workbook.xml'].async('string')
  }

  static extractFromArchive(archive, file) {
    return archive.files[file].async('string')
  }

  static extractFileContent(file) {
    const zip = new JSZip();
    return zip.loadAsync(file)
  }

  static getWorksheet(worksheetNumber: number) {
    let suffix = (worksheetNumber > 0) ? worksheetNumber : ''
    let worksheetPath = `ppt/embeddings/Microsoft_Excel_Worksheet${suffix}.xlsx`

    return (archive) => {
      return archive.files[worksheetPath].async('arraybuffer')
    }
  }

  static async zipCopy(sourceArchive, sourceFile, targetArchive, targetFile) {
    let archive = await sourceArchive
    let content = archive.files[sourceFile].async('nodebuffer')
    
    return targetArchive.file(targetFile, content)
  }
  
  static writeOutputFile(location, content) {
    fs.writeFile(location, content, function(err: { message: any }) {
      if(err) {
        throw new Error(`Error writing output: ${err.message}`)
      } else {
        console.log(`output: ${location}`)
      }
    })
  }
}