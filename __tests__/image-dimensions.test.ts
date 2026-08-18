import * as fs from 'fs';
import { imageDimensions } from '../src/helper/image-dimensions';

const mediaDir = `${__dirname}/media`;

describe('imageDimensions', () => {
  test('reads PNG dimensions from the repo media fixtures', () => {
    expect(
      imageDimensions(fs.readFileSync(`${mediaDir}/test.png`)),
    ).toStrictEqual({ width: 200, height: 129 });
    expect(
      imageDimensions(fs.readFileSync(`${mediaDir}/feather.png`)),
    ).toStrictEqual({ width: 190, height: 191 });
    expect(
      imageDimensions(fs.readFileSync(`${mediaDir}/Dàngerous Dinösaur.png`)),
    ).toStrictEqual({ width: 256, height: 256 });
  });

  test('reads JPEG dimensions from a SOF0 segment', () => {
    const jpeg = Buffer.from([
      0xff, 0xd8, // SOI
      0xff, 0xc0, // SOF0
      0x00, 0x11, // segment length (17)
      0x08, // precision
      0x00, 0x81, // height 129
      0x00, 0xc8, // width 200
      0x03, // components
      0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
    ]);
    expect(imageDimensions(jpeg)).toStrictEqual({ width: 200, height: 129 });
  });

  test('reads JPEG dimensions when SOF follows other segments', () => {
    const app0 = Buffer.from([
      0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00,
      0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
    ]);
    const sof = Buffer.from([
      0xff, 0xc2, 0x00, 0x0b, 0x08, 0x01, 0x00, 0x02, 0x00, 0x01, 0x11, 0x00,
    ]);
    const jpeg = Buffer.concat([Buffer.from([0xff, 0xd8]), app0, sof]);
    expect(imageDimensions(jpeg)).toStrictEqual({ width: 512, height: 256 });
  });

  test('reads GIF dimensions', () => {
    const gif = Buffer.concat([
      Buffer.from('GIF89a', 'latin1'),
      Buffer.from([0x40, 0x01, 0xf0, 0x00, 0x00, 0x00]),
    ]);
    expect(imageDimensions(gif)).toStrictEqual({ width: 320, height: 240 });
  });

  test('reads BMP dimensions (40-byte DIB header, top-down negative height)', () => {
    const bmp = Buffer.alloc(40);
    bmp.write('BM', 0, 'latin1');
    bmp.writeUInt32LE(40, 14); // BITMAPINFOHEADER
    bmp.writeInt32LE(100, 18);
    bmp.writeInt32LE(-50, 22);
    expect(imageDimensions(bmp)).toStrictEqual({ width: 100, height: 50 });
  });

  test('reads BMP dimensions (12-byte BITMAPCOREHEADER)', () => {
    const bmp = Buffer.alloc(26);
    bmp.write('BM', 0, 'latin1');
    bmp.writeUInt32LE(12, 14);
    bmp.writeUInt16LE(64, 18);
    bmp.writeUInt16LE(32, 20);
    expect(imageDimensions(bmp)).toStrictEqual({ width: 64, height: 32 });
  });

  test('reads WebP lossless (VP8L) dimensions', () => {
    const webp = Buffer.alloc(30);
    webp.write('RIFF', 0, 'latin1');
    webp.write('WEBP', 8, 'latin1');
    webp.write('VP8L', 12, 'latin1');
    webp[20] = 0x2f;
    // 14-bit fields store size - 1: width 200, height 129
    webp.writeUInt32LE((200 - 1) | ((129 - 1) << 14), 21);
    expect(imageDimensions(webp)).toStrictEqual({ width: 200, height: 129 });
  });

  test('reads SVG dimensions from the repo media fixture', () => {
    expect(
      imageDimensions(fs.readFileSync(`${mediaDir}/test.svg`)),
    ).toStrictEqual({ width: 120, height: 120 });
  });

  test('reads SVG dimensions from viewBox and xml prolog variants', () => {
    expect(
      imageDimensions(Buffer.from('<svg viewBox="0 0 640 480"></svg>')),
    ).toStrictEqual({ width: 640, height: 480 });
    expect(
      imageDimensions(
        Buffer.from(
          '<?xml version="1.0" encoding="UTF-8"?>\n<svg width="10px" height="20px"></svg>',
        ),
      ),
    ).toStrictEqual({ width: 10, height: 20 });
  });

  test('throws on empty and too-short buffers', () => {
    expect(() => imageDimensions(Buffer.alloc(0))).toThrow(
      'Image buffer is empty or too short',
    );
    expect(() => imageDimensions(Buffer.from('GIF89a'))).toThrow(
      'Image buffer is empty or too short',
    );
  });

  test('throws on a truncated PNG', () => {
    const png = Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      Buffer.alloc(8),
    ]);
    expect(() => imageDimensions(png)).toThrow('Malformed PNG');
  });

  test('throws on a PNG without IHDR', () => {
    const png = Buffer.alloc(32);
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).copy(png);
    png.write('JUNK', 12, 'latin1');
    expect(() => imageDimensions(png)).toThrow('Malformed PNG');
  });

  test('throws (and returns promptly) on a JPEG with a zero-length segment', () => {
    // regression guard for the infinite-loop class of image-size advisories
    const jpeg = Buffer.from([
      0xff, 0xd8, 0xff, 0xe0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ]);
    expect(() => imageDimensions(jpeg)).toThrow(
      'Malformed JPEG: invalid segment length',
    );
  });

  test('throws on a JPEG that reaches SOS without a SOF segment', () => {
    const jpeg = Buffer.from([
      0xff, 0xd8, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x3f, 0x00, 0x00,
    ]);
    expect(() => imageDimensions(jpeg)).toThrow('Malformed JPEG: no size found');
  });

  test('throws on unrecognized formats (including former image-size formats)', () => {
    expect(() => imageDimensions(Buffer.alloc(20, 7))).toThrow(
      'Unsupported or unrecognized image format',
    );
    // ICNS magic — intentionally unsupported (vulnerable parser in image-size)
    const icns = Buffer.alloc(20);
    icns.write('icns', 0, 'latin1');
    expect(() => imageDimensions(icns)).toThrow(
      'Unsupported or unrecognized image format',
    );
  });
});
