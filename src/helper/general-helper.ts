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

export const vd = (v: any): void => {
  console.dir(v, { depth: 10 });
};
