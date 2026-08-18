/**
 * Minimal, loop-safe image dimension detection for the formats PowerPoint
 * accepts as slide media: PNG, JPEG, GIF, BMP, WebP and SVG. Every parser
 * either returns dimensions or throws — bounded iteration only, so
 * malformed or unsupported input can never hang the process.
 */

export interface ImageDimensions {
  width: number;
  height: number;
}

// Longest prefix any sniffer below needs to look at
const MIN_SNIFF_LENGTH = 12;
// SVGs are text; only the prolog + root tag are needed for dimensions
const SVG_SCAN_LIMIT = 4096;

const parsePng = (buffer: Buffer): ImageDimensions => {
  if (buffer.length < 24 || buffer.toString('latin1', 12, 16) !== 'IHDR') {
    throw new Error('Malformed PNG: missing IHDR chunk');
  }
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
};

const parseJpeg = (buffer: Buffer): ImageDimensions => {
  let offset = 2;
  while (offset + 1 < buffer.length) {
    if (buffer[offset] !== 0xff) {
      throw new Error('Malformed JPEG: expected marker');
    }
    // 0xFF fill bytes may pad the space between segments
    while (offset < buffer.length && buffer[offset] === 0xff) {
      offset++;
    }
    if (offset >= buffer.length) {
      break;
    }
    const marker = buffer[offset];
    offset++;

    // standalone markers without a length field
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      continue;
    }
    // dimensions always precede the scan data in a valid JPEG
    if (marker === 0xd9 || marker === 0xda) {
      break;
    }

    if (offset + 2 > buffer.length) {
      break;
    }
    const segmentLength = buffer.readUInt16BE(offset);
    if (segmentLength < 2) {
      throw new Error('Malformed JPEG: invalid segment length');
    }

    const isSof =
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== 0xc4 &&
      marker !== 0xc8 &&
      marker !== 0xcc;
    if (isSof) {
      if (offset + 7 > buffer.length) {
        break;
      }
      return {
        height: buffer.readUInt16BE(offset + 3),
        width: buffer.readUInt16BE(offset + 5),
      };
    }

    offset += segmentLength;
  }
  throw new Error('Malformed JPEG: no size found');
};

const parseGif = (buffer: Buffer): ImageDimensions => {
  if (buffer.length < 10) {
    throw new Error('Malformed GIF: buffer too short');
  }
  return {
    width: buffer.readUInt16LE(6),
    height: buffer.readUInt16LE(8),
  };
};

const parseBmp = (buffer: Buffer): ImageDimensions => {
  if (buffer.length < 26) {
    throw new Error('Malformed BMP: buffer too short');
  }
  const dibHeaderSize = buffer.readUInt32LE(14);
  if (dibHeaderSize === 12) {
    // legacy BITMAPCOREHEADER
    return {
      width: buffer.readUInt16LE(18),
      height: buffer.readUInt16LE(20),
    };
  }
  return {
    width: buffer.readInt32LE(18),
    // top-down BMPs store a negative height
    height: Math.abs(buffer.readInt32LE(22)),
  };
};

const readUInt24LE = (buffer: Buffer, offset: number): number =>
  buffer[offset] | (buffer[offset + 1] << 8) | (buffer[offset + 2] << 16);

const parseWebp = (buffer: Buffer): ImageDimensions => {
  if (buffer.length < 30) {
    throw new Error('Malformed WebP: buffer too short');
  }
  const chunkType = buffer.toString('latin1', 12, 16);
  if (chunkType === 'VP8 ') {
    if (
      buffer[23] !== 0x9d ||
      buffer[24] !== 0x01 ||
      buffer[25] !== 0x2a
    ) {
      throw new Error('Malformed WebP: bad VP8 start code');
    }
    return {
      width: buffer.readUInt16LE(26) & 0x3fff,
      height: buffer.readUInt16LE(28) & 0x3fff,
    };
  }
  if (chunkType === 'VP8L') {
    if (buffer[20] !== 0x2f) {
      throw new Error('Malformed WebP: bad VP8L signature');
    }
    const bits = buffer.readUInt32LE(21);
    return {
      width: (bits & 0x3fff) + 1,
      height: ((bits >> 14) & 0x3fff) + 1,
    };
  }
  if (chunkType === 'VP8X') {
    return {
      width: readUInt24LE(buffer, 24) + 1,
      height: readUInt24LE(buffer, 27) + 1,
    };
  }
  throw new Error('Malformed WebP: unknown chunk type');
};

const svgAttribute = (tag: string, name: string): number | undefined => {
  const match = tag.match(
    new RegExp('[\\s"\']' + name + '\\s*=\\s*["\']([0-9.]+)(?:px)?["\']'),
  );
  return match ? Math.round(parseFloat(match[1])) : undefined;
};

const parseSvg = (buffer: Buffer): ImageDimensions => {
  const text = buffer.toString('utf8', 0, SVG_SCAN_LIMIT).replace(/^\uFEFF/, '');
  const svgStart = text.indexOf('<svg');
  const tagEnd = svgStart === -1 ? -1 : text.indexOf('>', svgStart);
  if (svgStart === -1 || tagEnd === -1) {
    throw new Error('Malformed SVG: no <svg> root tag found');
  }
  const tag = text.slice(svgStart, tagEnd + 1);

  const width = svgAttribute(tag, 'width');
  const height = svgAttribute(tag, 'height');
  if (width !== undefined && height !== undefined) {
    return { width, height };
  }

  const viewBox = tag.match(
    /viewBox\s*=\s*["']\s*[0-9.+-]+[\s,]+[0-9.+-]+[\s,]+([0-9.]+)[\s,]+([0-9.]+)\s*["']/,
  );
  if (viewBox) {
    return {
      width: Math.round(parseFloat(viewBox[1])),
      height: Math.round(parseFloat(viewBox[2])),
    };
  }
  throw new Error('Malformed SVG: no width/height or viewBox found');
};

// binary formats are sniffed first, so any <svg tag within the scanned
// prefix (after an optional prolog, DOCTYPE, comments or a generator
// banner) is a safe signal; parseSvg still throws on malformed content
const looksLikeSvg = (buffer: Buffer): boolean =>
  buffer
    .toString('utf8', 0, Math.min(buffer.length, SVG_SCAN_LIMIT))
    .includes('<svg');

/**
 * Detects the pixel dimensions of an image buffer.
 * Supported: PNG, JPEG, GIF, BMP, WebP, SVG — the formats PowerPoint
 * accepts as slide media. Throws on any other or malformed input.
 */
export const imageDimensions = (buffer: Buffer): ImageDimensions => {
  if (!buffer || buffer.length < MIN_SNIFF_LENGTH) {
    throw new Error('Image buffer is empty or too short');
  }

  if (buffer.readUInt32BE(0) === 0x89504e47 && buffer.readUInt32BE(4) === 0x0d0a1a0a) {
    return parsePng(buffer);
  }
  if (buffer[0] === 0xff && buffer[1] === 0xd8) {
    return parseJpeg(buffer);
  }
  const ascii6 = buffer.toString('latin1', 0, 6);
  if (ascii6 === 'GIF87a' || ascii6 === 'GIF89a') {
    return parseGif(buffer);
  }
  if (buffer[0] === 0x42 && buffer[1] === 0x4d) {
    return parseBmp(buffer);
  }
  if (
    buffer.toString('latin1', 0, 4) === 'RIFF' &&
    buffer.toString('latin1', 8, 12) === 'WEBP'
  ) {
    return parseWebp(buffer);
  }
  if (looksLikeSvg(buffer)) {
    return parseSvg(buffer);
  }

  throw new Error('Unsupported or unrecognized image format');
};
