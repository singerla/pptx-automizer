import { Workbook } from "../definitions/app"
import XmlHelper from "./xml"

export const setSolidFill = (element) => {
  element.getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6')
}

export const setText = (text: string) => (element) => {
  element.getElementsByTagName('a:t')[0]
    .firstChild
    .data = text
}

export const revertElements = (doc: Document) => {
  // console.log(doc)
}

export const setPosition = (pos: any) => (element: HTMLElement) => {
  console.log(element.getElementsByTagName('p:cNvPr')[0].getAttribute('name'))

  element.getElementsByTagName('a:off')[0].setAttribute('x', pos.x)
}

export const setChartData = (data: any) => (element: HTMLElement, chart: Document, workbook: Workbook) => {
  XmlHelper.setChartData(chart, workbook, data)
}