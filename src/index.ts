import Automizer from "../src/automizer"
import FileHelper from "./helper/file"

let textMod = (slide: any) => {
  slide.getElementsByTagName('p:sp')[1]
    .getElementsByTagName('a:t')[0]
    .firstChild
    .data = 'Test 123 2'
}

async function main() {
  const automize = new Automizer
  
  let pres = automize.importRootTemplate('pptx/Presentation.pptx')
    .importTemplate('pptx/Graph-Test.pptx', 'graphTpl')
    .importTemplate('pptx/Graph-Test2.pptx', 'graphTpl2')
    .importTemplate('pptx/Graph-Test3.pptx', 'graphTpl3')
    .importTemplate('pptx/Slide-Test.pptx', 'slideTpl')
    .importTemplate('pptx/Slide-Test3.pptx', 'arrows')
  // .importTemplate('pptx/Slide-Test2.pptx', 'slideTpl2')
  
  pres.addSlide('slideTpl', 1)
    .modify(textMod)
    
  for(let i=0; i<=10; i++) {
    pres.addSlide('arrows', 1)
  }


  pres.addSlide('graphTpl', 1)
  
  for(let i=0; i<=10; i++) {
    pres.addSlide('arrows', 2)
    pres.addSlide('graphTpl3', 2)
  }

  pres.addSlide('graphTpl2', 1)

  pres.addSlide('graphTpl3', 2)

  pres.write('out/Test-out.pptx')
}

main()

FileHelper.extractAllForecefully('out/Test-out.pptx')