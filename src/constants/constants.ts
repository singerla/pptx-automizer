import {
  TargetByRelIdMapParam,
  TrackedRelation,
  TrackedRelationTag,
} from '../types/types';

export const TargetByRelIdMap: Record<string, TargetByRelIdMapParam> = {
  chart: {
    relRootTag: 'c:chart',
    relAttribute: 'r:id',
    prefix: '../charts/chart',
  },
  chartEx: {
    relRootTag: 'cx:chart',
    relAttribute: 'r:id',
    prefix: '../charts/chartEx',
  },
  image: {
    relRootTag: 'a:blip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
  },
  'image:svg': {
    relRootTag: 'asvg:svgBlip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
  },
  'image:media': {
    relRootTag: 'p14:media',
    relAttribute: 'r:embed',
    relType: 'http://schemas.microsoft.com/office/2007/relationships/media',
    prefix: '../media/media',
    findAll: true,
  },
  'image:audioFile': {
    relRootTag: 'a:audioFile',
    relAttribute: 'r:link',
    prefix: '../media/media',
    relType:
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
    findAll: true,
  },
  'image:videoFile': {
    relRootTag: 'a:videoFile',
    relAttribute: 'r:link',
    prefix: '../media/media',
    relType:
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    findAll: true,
  },
  hyperlink: {
    relRootTag: 'a:hlinkClick',
    relAttribute: 'r:id',
    prefix: '',
    findAll: true,
  },
  oleObject: {
    relRootTag: 'p:oleObj',
    relAttribute: 'r:id',
    prefix: '../embeddings/oleObject',
  },
};

export const imagesTrack: () => TrackedRelation[] = () => [
  {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    tag: 'a:blip',
    role: 'image',
    attribute: 'r:embed',
  },
  {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    tag: 'asvg:svgBlip',
    role: 'image',
    attribute: 'r:embed',
  },
];

export const hyperlinksTrack: () => TrackedRelation[] = () => [
  {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    tag: 'a:hlinkClick',
    role: 'hyperlink',
    attribute: 'r:id',
  },
];

export const contentTrack = (): TrackedRelationTag[] => {
  return [
    {
      source: 'ppt/presentation.xml',
      relationsKey: 'ppt/_rels/presentation.xml.rels',
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
          tag: 'p:sldMasterId',
          role: 'slideMaster',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
          tag: 'p:sldId',
          role: 'slide',
        },
      ],
    },
    {
      source: 'ppt/slides',
      relationsKey: 'ppt/slides/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
          tag: 'c:chart',
          role: 'chart',
        },
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          role: 'slideLayout',
          tag: null,
        },
        ...imagesTrack(),
        ...hyperlinksTrack(),
      ],
    },
    {
      source: 'ppt/charts',
      relationsKey: 'ppt/charts/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
          tag: 'c:externalData',
          role: 'externalData',
        },
      ],
    },
    {
      source: 'ppt/slideMasters',
      relationsKey: 'ppt/slideMasters/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          tag: 'p:sldLayoutId',
          role: 'slideLayout',
        },
        ...imagesTrack(),
        ...hyperlinksTrack(),
      ],
    },
    {
      source: 'ppt/slideLayouts',
      relationsKey: 'ppt/slideLayouts/_rels',
      isDir: true,
      tags: [
        {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
          role: 'slideMaster',
          tag: null,
        },
        ...imagesTrack(),
        ...hyperlinksTrack(),
      ],
    },
  ];
};
