import Automizer from "../src/automizer"

async function main() {
  const automize = new Automizer
  
  
  // automize.importTemplate('pptx/Graph-Test.pptx', 'graphTpl')

  automize.importTemplate('pptx/Slide-Test.pptx', 'slideTpl')
  // automize.importTemplate('pptx/Slide-Test2.pptx', 'slideTpl2')
  automize.importTemplate('pptx/Slide-Test3.pptx', 'arrows')
  
  let pres = automize.importRootTemplate('pptx/Presentation.pptx')

  let slide = pres.addSlide('slideTpl', 1)

  slide.modify(slide => {
    slide.getElementsByTagName('p:sp')[1]
      .getElementsByTagName('a:t')[0]
      .firstChild
      .data = 'Test 123 2'
  })
  
  let slide2 = pres.addSlide('slideTpl', 1)
  slide2.modify(slide => {
    slide.getElementsByTagName('p:sp')[0]
      .getElementsByTagName('a:t')[0]
      .firstChild
      .data = 'mod2'
  })

  pres.addSlide('arrows', 2)
  pres.addSlide('arrows', 1)

  // slide.addChart('graphTpl', 0)

  pres.write('out/Test-out.pptx')
}

main()
