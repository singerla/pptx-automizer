export default class HyperlinkElement {
  private doc: Document;
  private relId: string;
  private isInternal: boolean;

  constructor(doc: Document, relId: string, isInternal: boolean) {
    this.doc = doc;
    this.relId = relId;
    this.isInternal = isInternal;
  }

  createHlinkClick(): Element {
    const hlinkClick = this.doc.createElement('a:hlinkClick');
    hlinkClick.setAttribute('r:id', this.relId);
    hlinkClick.setAttribute(
      'xmlns:r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    );

    if (this.isInternal) {
      hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
      hlinkClick.setAttribute(
        'xmlns:a',
        'http://schemas.openxmlformats.org/drawingml/2006/main',
      );
      hlinkClick.setAttribute(
        'xmlns:p14',
        'http://schemas.microsoft.com/office/powerpoint/2010/main',
      );
    }

    return hlinkClick;
  }

  createTextRun(text: string): Element {
    const run = this.doc.createElement('a:r');
    const rPr = this.doc.createElement('a:rPr');
    const t = this.doc.createElement('a:t');

    rPr.appendChild(this.createHlinkClick());
    t.textContent = text;

    run.appendChild(rPr);
    run.appendChild(t);

    return run;
  }
}
