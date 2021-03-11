import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import {
	IChart, Target
} from './types'


export default class Chart implements IChart {  
  sourceArchive: JSZip
  targetArchive: JSZip
  sourceNumber: number
  targetNumber: number
  sourceWorksheet: number | string
  targetWorksheet: number | string
  targetSlideNumber: number
  sourceRid: any

  constructor(relsXmlInfo: Target, sourceArchive: JSZip, targetSlideNumber: number) {
    this.sourceNumber = relsXmlInfo.number
    this.sourceRid = relsXmlInfo.rId
    this.sourceArchive = sourceArchive
    this.targetSlideNumber = targetSlideNumber
  }

  setTarget(archive: JSZip, number: number) {
    this.targetArchive = archive
    this.targetNumber = number
  }

  async append() {
    await this.copyFiles()
    await this.appendTypes()
  }

  async copyFiles() {
    this.copyChartFiles()

    let wbRelsPath = `ppt/charts/_rels/chart${this.sourceNumber}.xml.rels`
    let worksheets = await XmlHelper.getTargetsFromRelationships(this.sourceArchive, wbRelsPath, '../embeddings/Microsoft_Excel_Worksheet', '.xlsx')
    let worksheet = worksheets[0]

    this.sourceWorksheet = (worksheet.number === 0) ? '' : worksheet.number
    this.targetWorksheet = this.targetNumber
    
    this.copyWorksheetFile()
    this.editTargetWorksheetRel()
    this.editTargetChartRel()
  }

  async appendTypes() {
    await this.appendChartExtensionToContentType()
    await this.appendChartToContentType()
    await this.appendColorToContentType()
    await this.appendStyleToContentType()
  }

  async copyChartFiles(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/charts/chart${this.sourceNumber}.xml`, 
      this.targetArchive, `ppt/charts/chart${this.targetNumber}.xml`
    )

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/charts/colors${this.sourceNumber}.xml`, 
      this.targetArchive, `ppt/charts/colors${this.targetNumber}.xml`
    )

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/charts/style${this.sourceNumber}.xml`, 
      this.targetArchive, `ppt/charts/style${this.targetNumber}.xml`
    )

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/charts/_rels/chart${this.sourceNumber}.xml.rels`, 
      this.targetArchive, `ppt/charts/_rels/chart${this.targetNumber}.xml.rels`
    )
  }

  async editTargetChartRel() {
    let targetRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`
    let relXml = await XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile)
    let relations = relXml.getElementsByTagName('Relationship')

    for(let i in relations) {
      let element = relations[i]
      if(element.getAttribute) {
        let rId = element.getAttribute('Id')
        if(rId === this.sourceRid) {
          element.setAttribute('Target', `../charts/chart${this.targetNumber}.xml`)
        }
      }
    }

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml)
  }

  async editTargetWorksheetRel() {
    let targetRelFile = `ppt/charts/_rels/chart${this.targetNumber}.xml.rels`
    let relXml = await XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile)
    let relations = relXml.getElementsByTagName('Relationship')

    for(let i in relations) {
      let element = relations[i]
      if(element.getAttribute) {
        let type = element.getAttribute('Type')
        switch(type) {
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package":
            element.setAttribute('Target', `../embeddings/Microsoft_Excel_Worksheet${this.targetWorksheet}.xlsx`)
            break
          case "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle":
            element.setAttribute('Target', `colors${this.targetNumber}.xml`)
            break
          case "http://schemas.microsoft.com/office/2011/relationships/chartStyle":
            element.setAttribute('Target', `style${this.targetNumber}.xml`)
            break
        }
      }
    }
    
    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml)
  }

  async copyWorksheetFile(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/embeddings/Microsoft_Excel_Worksheet${this.sourceWorksheet}.xlsx`, 
      this.targetArchive, `ppt/embeddings/Microsoft_Excel_Worksheet${this.targetWorksheet}.xlsx`,
    )
  }

  appendChartExtensionToContentType(): Promise<HTMLElement | boolean> {
    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(this.targetArchive, {
        Extension: `xlsx`,
        ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
      }),
      tag: 'Default',
      clause: (xml: HTMLElement) => !XmlHelper.findByAttribute(xml, 'Default', 'Extension', 'xlsx')
    })
  }

  appendChartToContentType(): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/chart${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`
      })
    )
  }

  appendColorToContentType(): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/colors${this.targetNumber}.xml`,
        ContentType: `application/vnd.ms-office.chartcolorstyle+xml`
      })
    )
  }

  appendStyleToContentType(): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/style${this.targetNumber}.xml`,
        ContentType: `application/vnd.ms-office.chartstyle+xml`
      })
    )
  }

}
