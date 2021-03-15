import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import {
	IChart, RelationshipAttribute, RootPresTemplate, Target
} from './types'


export default class Chart implements IChart {  
  sourceArchive: JSZip
  targetArchive: JSZip
  sourceNumber: number
  targetNumber: number
  sourceWorksheet: number | string
  targetWorksheet: number | string
  targetTemplate: RootPresTemplate
  targetSlideNumber: number
  sourceRid: string
  appendMode: boolean
  createdRid: string

  constructor(relsXmlInfo: Target, sourceArchive: JSZip) {
    this.sourceNumber = relsXmlInfo.number
    this.sourceRid = relsXmlInfo.rId
    this.sourceArchive = sourceArchive
  }
  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number, appendMode?: boolean): Promise<void> {
    this.targetTemplate = targetTemplate
    this.targetArchive = await this.targetTemplate.archive
    this.targetNumber = this.targetTemplate.incrementCounter('charts')
    this.targetSlideNumber = targetSlideNumber

    if(appendMode !== undefined) {
      this.appendMode = appendMode
    }

    await this.copyFiles()
    await this.appendTypes()
  }

  async copyFiles(): Promise<void> {
    this.copyChartFiles()

    let wbRelsPath = `ppt/charts/_rels/chart${this.sourceNumber}.xml.rels`
    let worksheets = await XmlHelper.getTargetsFromRelationships(this.sourceArchive, wbRelsPath, '../embeddings/Microsoft_Excel_Worksheet', '.xlsx')
    let worksheet = worksheets[0]

    this.sourceWorksheet = (worksheet.number === 0) ? '' : worksheet.number
    this.targetWorksheet = this.targetNumber
    
    this.copyWorksheetFile()
    this.editTargetWorksheetRel()
    
    if(this.appendMode === true) {
      await this.appendToSlideRels()
    } else {
      await this.editTargetChartRel()
    }
  }

  async appendTypes(): Promise<void> {
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

  async editTargetChartRel(): Promise<void> {
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

  async appendToSlideRels(): Promise<HTMLElement> {
    let targetRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`
    this.createdRid = await XmlHelper.getNextRelId(this.targetArchive, targetRelFile)
    let attributes = <RelationshipAttribute> {
      Id: this.createdRid,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      Target: `../charts/chart${this.targetNumber}.xml`
    }

    return XmlHelper.append(
      XmlHelper.createRelationshipChild(this.targetArchive, targetRelFile, attributes)
    )
  }

  async editTargetWorksheetRel(): Promise<void>  {
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
