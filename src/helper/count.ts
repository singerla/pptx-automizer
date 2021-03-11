import JSZip from 'jszip'
import XmlHelper from './xml'

export default class CountHelper {

  static async countImages(presentation: JSZip): Promise<number> {
    let files = await presentation.file(/ppt\/media\/image/)
    return files.length
  }

  static async countSlides(presentation: JSZip): Promise<number> {
    let presentationXml = await XmlHelper.getXmlFromArchive(presentation, 'ppt/presentation.xml')
    let slideCount = presentationXml.getElementsByTagName('p:sldId').length
    return slideCount
  }

  static async countCharts(presentation: JSZip): Promise<number> {
    let contentTypesXml = await XmlHelper.getXmlFromArchive(presentation, '[Content_Types].xml')
    let overrides = contentTypesXml.getElementsByTagName('Override')
    let chartCount = 0

    for(let i in overrides) {
      let override = overrides[i]
      if(override.getAttribute) {
        let contentType = override.getAttribute('ContentType')
        if(contentType === `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`) {
          chartCount++
        }
      }
    }

    return chartCount
  }

}