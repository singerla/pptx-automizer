import Automizer from "./index"
import Slide from "./slide"
import { setChartData, setPosition } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithImages.pptx`, 'images')
  .load(`SlideWithLink.pptx`, 'link')
  .load(`SlideWithCharts.pptx`, 'charts')
  .load(`EmptySlide.pptx`, 'empty')

pres
  // .addSlide('images', 2)
  .addSlide('empty', 1, (slide: Slide) => {
    slide.addElement('charts', 2, 'PieChart', [
      setChartData({
        series: [
          { label: 'فارسی 1' },
        ],
        categories: [
          { label: 'নাগরিক', values: [ 12.5 ] },
          { label: 'cat 2', values: [ 14 ] },
          { label: 'cat 3', values: [ 15 ] },
          { label: 'cat 4', values: [ 26 ] }
        ]
      }), 
      setPosition({x: 8000000}) 
    ])

    slide.addElement('charts', 1, 'StackedBars', [
      setChartData({
        series: [
          { label: 'series 1' },
        ],
        categories: [
          { label: 'cat 2-1', values: [ 50 ] },
          { label: 'cat 2-2', values: [ 14 ] },
          { label: 'cat 2-3', values: [ 15 ] },
          { label: 'cat 2-4', values: [ 26 ] }
        ]
      }), 
      setPosition({x: 8000000})
    ])

    // slide.addElement('charts', 2, 'PieChart')
    // slide.addElement('charts', 2, 'PieChart')
    slide.addElement('images', 2, 'imageSVG', setPosition({x: 8000000}))
    // slide.addElement('link', 1, 'Link')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('charts', 1, 'StackedBars')
  })

  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
