/**
 * Child script for memory-bounded-large-deck.test.ts — runs in its own node
 * process with --expose-gc so heap numbers are deterministic. Builds a
 * synthetic Think-cell-like template (600 shapes per slide, each carrying a
 * p:custDataLst, ~300 KB slide XML), appends its slide N times, writes the
 * deck and prints a JSON result line to stdout.
 *
 * Not a test file itself: invoked by the test via
 * `node -r ts-node/register/transpile-only --expose-gc <this file> <slides>`.
 */
import JSZip from 'jszip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import Automizer from '../../src/index';

const slideCount = Number(process.argv[2] ?? 60);
const templateDir = path.join(__dirname, '..', 'pptx-templates');

function syntheticShape(i: number): string {
  const x = 100000 + (i % 30) * 300000;
  const y = 100000 + Math.floor(i / 30) * 300000;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${1000 + i}" name="Synthetic shape ${i}"/>` +
    `<p:cNvSpPr/><p:nvPr><p:custDataLst/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="250000" cy="250000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/>` +
    `<a:t>Sample text for synthetic shape number ${i} with some padding characters</a:t>` +
    `</a:r></a:p></p:txBody></p:sp>`
  );
}

async function buildSyntheticTemplate(workDir: string): Promise<string> {
  const source = await fs.promises.readFile(
    path.join(templateDir, 'EmptySlide.pptx'),
  );
  const zip = await new JSZip().loadAsync(source);
  let slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  const shapes = Array.from({ length: 600 }, (_, i) => syntheticShape(i));
  slideXml = slideXml.replace('</p:spTree>', shapes.join('') + '</p:spTree>');
  zip.file('ppt/slides/slide1.xml', slideXml);

  const target = path.join(workDir, 'SyntheticLargeDeck.pptx');
  await fs.promises.writeFile(
    target,
    await zip.generateAsync({ type: 'nodebuffer' }),
  );
  return target;
}

async function main() {
  const workDir = await fs.promises.mkdtemp(
    path.join(os.tmpdir(), 'pptx-automizer-memtest-'),
  );

  try {
    const synthetic = await buildSyntheticTemplate(workDir);

    const pres = new Automizer({
      templateDir,
      outputDir: workDir,
      removeExistingSlides: true,
    })
      .loadRoot('RootTemplate.pptx')
      .load(synthetic, 'big');

    for (let i = 0; i < slideCount; i++) {
      pres.addSlide('big', 1);
    }

    const summary = await pres.write('memory-bounded-output.pptx');

    const archive = (
      pres as unknown as {
        rootTemplate: { archive: { buffer: Map<string, unknown> } };
      }
    ).rootTemplate.archive;

    global.gc();
    const heapUsedMb = process.memoryUsage().heapUsed / 1024 / 1024;

    console.log(
      JSON.stringify({
        slides: summary.slides,
        status: summary.status,
        bufferedParts: archive.buffer.size,
        heapUsedMb: Math.round(heapUsedMb),
      }),
    );
  } finally {
    await fs.promises.rm(workDir, { recursive: true, force: true });
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
