export enum ContentTypeMap {
  jpg = 'image/jpeg',
  jpeg = 'image/jpeg',
  png = 'image/png',
  gif = 'image/gif',
  svg = 'image/svg+xml',
  mp3 = 'audio/mp3',
  m4v = 'video/mp4',
  mp4 = 'video/mp4',
  emf = 'image/x-emf',
  wdp = 'image/vnd.ms-photo',

  // This is required to support think-cell contents
  xml = 'application/xml',
  bin = 'application/vnd.openxmlformats-officedocument.oleObject',
  vml = 'application/vnd.openxmlformats-officedocument.vmlDrawing',
}

export type ContentTypeExtension = keyof typeof ContentTypeMap;
