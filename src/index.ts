import Automizer from "../src/automizer"

async function main() {
  const automize = new Automizer
  
  let pres = automize.importRootTemplate('pptx/Presentation.pptx')
  
  automize.importTemplate('pptx/Graph-Test2.pptx', 'graphTpl')

  automize.importTemplate('pptx/Slide-Test.pptx', 'slideTpl')
  // automize.importTemplate('pptx/Slide-Test2.pptx', 'slideTpl2')
  automize.importTemplate('pptx/Slide-Test3.pptx', 'arrows')
  
  let slide = pres.addSlide('slideTpl', 1)

  // slide.modify(slide => {
  //   slide.getElementsByTagName('p:sp')[1]
  //     .getElementsByTagName('a:t')[0]
  //     .firstChild
  //     .data = 'Test 123 2'
  // })
    
  for(let i=0; i<=10; i++) {
    pres.addSlide('arrows', 2)
    pres.addSlide('arrows', 1)
  }


  pres.addSlide('graphTpl', 1)

  pres.write('out/Test-out.pptx')
}

main()
