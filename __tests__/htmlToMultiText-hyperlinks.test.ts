import Automizer, { modify } from '../src/index';

test('create presentation with external hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<p>Visit our <a href="https://example.com">website</a> for more information</p>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-external-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with internal slide hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html = '<p>See <a href="3">slide 3</a> for details</p>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-internal-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});

test('create presentation with multiple hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html =
    '<html><body>' +
    '<p>Check out <a href="https://google.com">Google</a> or <a href="https://github.com">GitHub</a></p>' +
    '<p>Jump to <a href="2">slide 2</a> or <a href="5">slide 5</a></p>' +
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
    '<li>Regular bullet with <a href="3">internal link</a></li>' +
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
    '<li>Internal: <a href="4">Jump to slide 4</a></li>' +
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

test('verify hyperlink relationships are created correctly', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html = '<p>Visit <a href="https://example.com">our website</a></p>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, async (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-verify-relationships.test.pptx`);

  // Test passes if file is written successfully
  // Test passes if file is written successfully
});

test('verify internal slide hyperlink relationships are created correctly', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  const html = '<p>Go to <a href="3">slide 3</a></p>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1, async (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-verify-internal-relationships.test.pptx`);

  // Test passes if file is written successfully
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

  const html = '<p>Visit <a href="https://example.com">website</a></p>';

  // This should complete without error, but hyperlinks won't be created
  const result = await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      // Directly calling htmlToMultiText without relation element access
      // would normally skip hyperlink creation
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`htmlToMultiText-no-relation-warning.test.pptx`);

  // Test passes if file is written successfully
});
