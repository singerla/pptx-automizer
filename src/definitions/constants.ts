import { TargetByRelIdMapParam } from './app';

export const TargetByRelIdMap = {
  chart: <TargetByRelIdMapParam>{
    relRootTag: 'c:chart',
    relAttribute: 'r:id',
    prefix: '../charts/chart'
  },
  image: <TargetByRelIdMapParam>{
    relRootTag: 'a:blip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
    expression: /\..+?$/
  },
  'image:svg': <TargetByRelIdMapParam>{
    relRootTag: 'asvg:svgBlip',
    relAttribute: 'r:embed',
    prefix: '../media/image',
    expression: /\..+?$/
  }
};
