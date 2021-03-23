// Thanks to https://github.com/aishwar/xml-pretty-print
// Alternative: https://github.com/riversun/xml-beautify

type PrettyPrintToken = {
  match: string;
  tag: string;
  offset: number;
  preContent: string;
};

export class XmlPrettyPrint {
  xmlStr: string;
  TAB: string;

  constructor(xmlStr: string) {
    this.xmlStr = xmlStr;
    this.TAB = '  ';
  }

  dump(): void {
    console.log(this.prettify());
  }

  prettify(): string {
    return this.parse(this.xmlStr).join('\n');
  }

  parse(xmlStr: string): string[] {
    const opener = /<(\w+)[^>]*?>/m;
    const closer = /<\/[^>]*>/m;
    let idx = 0;
    let indent = 0;
    let processing = '';
    const tags = [];
    const output = [];

    while (idx < xmlStr.length) {
      processing += xmlStr[idx];

      const openToken = this.getToken(opener, processing);
      const closeToken = this.getToken(closer, processing);

      if (openToken) {
        // Check if it is a singular element, e.g. <link />
        if (processing[processing.length - 2] != '/') {
          this.addLine(output, openToken.preContent, indent);
          this.addLine(output, openToken.match, indent);

          tags.push(openToken.tag);
          indent += 1;
          processing = '';
        } else {
          this.addLine(output, openToken.preContent, indent);
          this.addLine(output, openToken.match, indent);
          processing = '';
        }
      } else if (closeToken) {
        this.addLine(output, closeToken.preContent, indent);

        if (tags[tags.length] == closeToken.tag) {
          tags.pop();
          indent -= 1;
        }

        this.addLine(output, closeToken.match, indent);
        processing = '';
      }

      idx += 1;
    }

    if (tags.length) {
      console.log(
        'WARNING: xmlFile may be malformed. Not all opening tags were closed. Following tags were left open:',
      );
      console.log(tags);
    }

    return output;
  }

  getToken(regex: RegExp, str: string): PrettyPrintToken {
    if (regex.test(str)) {
      const matches = regex.exec(str);
      const match = matches[0];
      const offset = str.length - match.length;
      const preContent = str.substring(0, offset);

      return {
        match,
        tag: matches[1],
        offset,
        preContent,
      };
    }
  }

  addLine(output: string[], content: string, indent: number): void {
    // Trim the content
    content = content.replace(/^\s+|\s+$/, '');
    if (content) {
      let tabs = '';

      while (indent--) {
        tabs += this.TAB;
      }
      output.push(tabs + content);
    }
  }
}
