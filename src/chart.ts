import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import {
	IChart
} from './types/interfaces'

import {
	Target
} from './types/xml'


export default class Chart implements IChart {
  sourceNumber: number
  targetNumber: number

  sourceArchive: JSZip
  targetArchive: JSZip

  constructor(relsXmlInfo: Target, sourceArchive: JSZip) {
    this.sourceNumber = relsXmlInfo.number
    this.sourceArchive = sourceArchive
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

    let number = (worksheet.number === 0) ? '' : worksheet.number
    this.copyWorksheetFile(number, number)
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

  async copyWorksheetFile(sourceWorksheet: number | string, targetWorksheet: number|string): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/embeddings/Microsoft_Excel_Worksheet${sourceWorksheet}.xlsx`, 
      this.targetArchive, `ppt/embeddings/Microsoft_Excel_Worksheet${targetWorksheet}.xlsx`
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
