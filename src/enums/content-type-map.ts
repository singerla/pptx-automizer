export enum ContentTypeMap {
  jpg = 'image/jpeg',
  // eslint-disable-next-line @typescript-eslint/no-duplicate-enum-values -- extension alias
  jpeg = 'image/jpeg',
  png = 'image/png',
  gif = 'image/gif',
  svg = 'image/svg+xml',
  mp3 = 'audio/mp3',
  m4v = 'video/mp4',
  // eslint-disable-next-line @typescript-eslint/no-duplicate-enum-values -- extension alias
  mp4 = 'video/mp4',
  emf = 'image/x-emf',
  wdp = 'image/vnd.ms-photo',

  // This is required to support think-cell contents
  xml = 'application/xml',
  bin = 'application/vnd.openxmlformats-officedocument.oleObject',
  vml = 'application/vnd.openxmlformats-officedocument.vmlDrawing',
}

export type ContentTypeExtension = keyof typeof ContentTypeMap;
