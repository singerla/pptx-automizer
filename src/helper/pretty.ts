// Thanks to https://github.com/aishwar/xml-pretty-print
// Alternative: https://github.com/riversun/xml-beautify

export default class XmlPrettyPrint {
  xmlStr: string;
  TAB: string;

  constructor(xmlStr) {
    this.xmlStr = xmlStr;
    this.TAB = '  ';
  }

  dump() {
    console.log(this.prettify());
  }

  prettify() {
    return this.parse(this.xmlStr).join('\n');
  }

  parse(xmlStr) {
    const opener = /<(\w+)[^>]*?>/m;
    const closer = /<\/[^>]*>/m;

    let idx = 0;
    let indent = 0;
    let processing = '';
    const tags = [];
    const output = [];
    let token;

    while (idx < xmlStr.length) {
      processing += xmlStr[idx];

      if (token = this.getToken(opener, processing)) {
        // Check if it is a singular element, e.g. <link />
        if (processing[processing.length - 2] != '/') {
          this.addLine(output, token.preContent, indent);
          this.addLine(output, token.match, indent);

          tags.push(token.tag);
          indent += 1;
          processing = '';
        } else {
          this.addLine(output, token.preContent, indent);
          this.addLine(output, token.match, indent);
          processing = '';
        }
      } else if (token = this.getToken(closer, processing)) {
        this.addLine(output, token.preContent, indent);

        if (tags[tags.length] == token.tag) {
          tags.pop();
          indent -= 1;
        }

        this.addLine(output, token.match, indent);
        processing = '';
      }

      idx += 1;
    }

    if (tags.length) {
      console.log('WARNING: xmlFile may be malformed. Not all opening tags were closed. Following tags were left open:');
      console.log(tags);
    }

    return output;
  }

  getToken(regex, str) {
    if (regex.test(str)) {
      const matches = regex.exec(str);
      const match = matches[0];
      const offset = str.length - match.length;
      const preContent = str.substring(0, offset);

      return {
        match,
        tag: matches[1],
        offset,
        preContent
      };
    }
  }

  addLine(output, content, indent) {
    // Trim the content
    if (content = content.replace(/^\s+|\s+$/, '')) {
      let tabs = '';

      while (indent--) {
        tabs += this.TAB;
      }
      output.push(tabs + content);
    }
  }
}
