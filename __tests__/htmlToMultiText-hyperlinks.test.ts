import Automizer, { modify, XmlElement } from '../src/index';
import { HtmlToMultiTextHelper } from '../src/helper/html-to-multitext-helper';
import { ModifyTextHelper } from '../src';

test('create presentation with external hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<body><p>Visit our <a href="https://example.com">website</a> for more information</p></body>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-external-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with multiple hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  // Please make sure that the internally linked slide does exist, otherwise the link will not be active.
  const html =
    '<html><body>' +
    '<p>Check out <a href="https://google.com">Google</a> or <a href="https://github.com">GitHub</a></p>' +
    '<p>Jump to <a href="2">slide 2</a> or <a href="5">slide 5 not existing</a></p>' +
    '</body></html>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-multiple-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with hyperlinks and formatting using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<html><body>' +
    '<p><strong>Important:</strong> Visit <a href="https://example.com"><em>our site</em></a></p>' +
    '<p>Regular text with <a href="https://test.com"><strong>bold link</strong></a> inline</p>' +
    '</body></html>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-hyperlinks-with-formatting.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with hyperlinks in bullet lists using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<html><body>' +
    '<ul>' +
    '<li><a href="https://site1.com">First link</a></li>' +
    '<li><a href="https://site2.com">Second link</a></li>' +
    '<li>Regular bullet with <a href="1">internal link</a></li>' +
    '</ul>' +
    '</body></html>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-hyperlinks-in-lists.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with mixed external and internal hyperlinks', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<html><body>' +
    '<p>Visit <a href="https://example.com">example.com</a> or go to <a href="2">slide 2</a></p>' +
    '<ul>' +
    '<li>External: <a href="https://google.com">Google</a></li>' +
    '<li>Internal: <a href="1">Jump to slide 1</a></li>' +
    '<li>More info: <a href="https://github.com">GitHub</a></li>' +
    '</ul>' +
    '</body></html>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-mixed-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});

test('htmlToMultiText with hyperlinks but no relation element should log warning', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    verbosity: 2,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html = '<body><p>Visit <a href="https://example.com">website</a></p></body>';

  // This should complete without error, but hyperlinks won't be created
  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      // Directly calling htmlToMultiText without relation element access
      // would normally skip hyperlink creation
      // A repair message is displayed on opening in PowerPoint.
      const paragraphs = new HtmlToMultiTextHelper().run(html);
      return (element: XmlElement): void => {
        ModifyTextHelper.setMultiText(paragraphs)(element);
      };
    })
    .write(`htmlToMultiText-no-relation-warning.test.pptx`);

  // Test passes if file is written successfully
});
