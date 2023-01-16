export class GeneralHelper {
  static arrayify<T>(input: T | T[]): T[] {
    if (Array.isArray(input)) {
      return input;
    } else if (input !== undefined) {
      return [input];
    } else {
      return [];
    }
  }

  static propertyExists<T>(object: T, property: string): boolean {
    if (!object || typeof object !== 'object') return false;
    return !!Object.getOwnPropertyDescriptor(object, property);
  }
}

export const vd = (v: any, keys?: boolean): void => {
  if (keys && typeof v === 'object') {
    v = Object.keys(v);
  }
  console.log('--------- [pptx-automizer] ---------');
  // @ts-ignore
  console.log(new Error().stack.split('\n')[2].trim());
  console.dir(v, { depth: 10 });
};
