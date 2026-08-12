import fs from 'fs';
import JSZip from 'jszip';
import {
  DOMParser,
  Document as XmlDocument,
  Element as XmlElement,
} from '@xmldom/xmldom';

/**
 * ROADMAP Phase 5, Tier 0 — output-assertion helper.
 *
 * Opens a written .pptx from __tests__/pptx-output and asserts on the XML of
 * a single part:
 *
 *   const slide = await expectXml('my-output.test.pptx', 'ppt/slides/slide1.xml');
 *   slide.toContainElement('a:t', 'my text');
 *   slide.toContainElementTimes('a:p', 3);
 *   slide.toHaveAttribute('c:idx', 'val', '2');
 *
 * For anything beyond these, `doc()` exposes the parsed part and `raw()` the
 * XML string, so a suite can drop down to expect() on its own.
 */

const outputDir = `${__dirname}/../pptx-output`;

export async function expectXml(
  outputFile: string,
  partPath: string,
): Promise<XmlPartExpectation> {
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/${outputFile}`),
  );
  const part = archive.file(partPath);
  if (!part) {
    throw new Error(`${outputFile} has no part ${partPath}`);
  }
  const xml = await part.async('text');
  return new XmlPartExpectation(outputFile, partPath, xml);
}

export class XmlPartExpectation {
  private readonly document: XmlDocument;

  constructor(
    private readonly outputFile: string,
    private readonly partPath: string,
    private readonly xml: string,
  ) {
    this.document = new DOMParser().parseFromString(xml, 'application/xml');
  }

  raw(): string {
    return this.xml;
  }

  doc(): XmlDocument {
    return this.document;
  }

  elements(tagName: string): XmlElement[] {
    const collection = this.document.getElementsByTagName(tagName);
    const result: XmlElement[] = [];
    for (let i = 0; i < collection.length; i++) {
      result.push(collection[i]);
    }
    return result;
  }

  /**
   * Asserts that at least one `tagName` element exists — and, when `text` is
   * given, that one of them has exactly this text content.
   */
  toContainElement(tagName: string, text?: string): this {
    const elements = this.elements(tagName);
    if (elements.length === 0) {
      throw new Error(`${this.describe()} contains no <${tagName}>`);
    }
    if (text !== undefined) {
      const texts = elements.map((element) => element.textContent);
      if (!texts.includes(text)) {
        throw new Error(
          `${this.describe()} has no <${tagName}> with text "${text}" ` +
            `(found: ${texts.map((value) => `"${value}"`).join(', ')})`,
        );
      }
    }
    return this;
  }

  toContainElementTimes(tagName: string, count: number): this {
    const found = this.elements(tagName).length;
    if (found !== count) {
      throw new Error(
        `${this.describe()} contains <${tagName}> ${found} times, ` +
          `expected ${count}`,
      );
    }
    return this;
  }

  /**
   * Asserts that some `tagName` element carries `attribute`, and, when
   * `value` is given, that one of them has exactly this value.
   */
  toHaveAttribute(tagName: string, attribute: string, value?: string): this {
    const carriers = this.elements(tagName).filter((element) =>
      element.hasAttribute(attribute),
    );
    if (carriers.length === 0) {
      throw new Error(
        `${this.describe()} has no <${tagName}> with attribute ${attribute}`,
      );
    }
    if (value !== undefined) {
      const values = carriers.map((element) => element.getAttribute(attribute));
      if (!values.includes(value)) {
        throw new Error(
          `${this.describe()} has no <${tagName} ${attribute}="${value}"> ` +
            `(found: ${values.map((found) => `"${found}"`).join(', ')})`,
        );
      }
    }
    return this;
  }

  private describe(): string {
    return `${this.outputFile} → ${this.partPath}`;
  }
}
