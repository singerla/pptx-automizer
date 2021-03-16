
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