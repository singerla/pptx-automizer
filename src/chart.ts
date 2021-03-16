import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import Shape from './shape'

import { IChart, RootPresTemplate, Target } from './types/app'
import { RelationshipAttribute } from './types/xml'

export default class Chart extends Shape implements IChart {  

  sourceWorksheet: number | string
  targetWorksheet: number | string

  constructor(relsXmlInfo: Target, sourceArchive: JSZip, sourceSlideNumber?:number) {
    super(relsXmlInfo, sourceArchive, sourceSlideNumber)

    this.relRootTag = 'c:chart'
    this.relAttribute = 'r:id'
    this.relParent = element => <HTMLElement> element.parentNode.parentNode.parentNode
  }
  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number, appendToTree?: boolean): Promise<Chart> {
    await this.setTarget(targetTemplate, targetSlideNumber)

    this.targetNumber = this.targetTemplate.incrementCounter('charts')

    await this.copyFiles()
    await this.appendTypes()
    await this.appendToSlideRels()

    if(appendToTree) {
      await this.appendToSlideTree()
    }

    await this.updateElementRelId()

    return this
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

  async appendToSlideRels(): Promise<HTMLElement> {
    this.createdRid = await XmlHelper.getNextRelId(this.targetArchive, this.targetSlideRelFile)
    let attributes = <RelationshipAttribute> {
      Id: this.createdRid,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      Target: `../charts/chart${this.targetNumber}.xml`
    }

    return XmlHelper.append(
      XmlHelper.createRelationshipChild(this.targetArchive, this.targetSlideRelFile, attributes)
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

  static async getAllOnSlide(archive: JSZip, relsPath: string): Promise<Target[]> {
    return await XmlHelper.getTargetsFromRelationships(archive, relsPath, '../charts/chart')
  }
}
