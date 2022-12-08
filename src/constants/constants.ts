import { TargetByRelIdMapParam } from '../types/types';

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
