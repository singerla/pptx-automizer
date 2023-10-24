import {
  TargetByRelIdMapParam,
  TrackedRelation,
  TrackedRelationTag,
} from '../types/types';

export const TargetByRelIdMap = {
  chart: {
    relRootTag: 'c:chart',
    relAttribute: 'r:id',
    prefix: '../charts/chart',
  } as TargetByRelIdMapParam,
  chartEx: {
    relRootTag: 'cx:chart',
    relAttribute: 'r:id',
    prefix: '../charts/chartEx',
  } as TargetByRelIdMapParam,
  image: {
    relRootTag: 'a:blip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
  } as TargetByRelIdMapParam,
  'image:svg': {
    relRootTag: 'asvg:svgBlip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
  } as TargetByRelIdMapParam,
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
      ],
    },
  ];
};
