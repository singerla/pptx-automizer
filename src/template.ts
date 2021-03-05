import {
 PresTemplate
} from './types/interfaces'

import FileHelper from './helper/file'

class Template {

  static import(location: string, name?:string): PresTemplate {
    let file = FileHelper.readFile(location)
    let archive = FileHelper.extractFileContent(file)
    
    let newTemplate = <PresTemplate> <unknown>{
      location: location,
      file: file,
      archive: archive,
      slides: []
    }
    
    if(name) {
      newTemplate.name = name
    }
    
    return newTemplate
  }

}


export default Template