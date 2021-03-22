export class GeneralHelper {
  static arrayify<T>(input: T): T[] {
    if (input instanceof Array) {
      return input;
    } else if (input !== undefined) {
      return [input];
    } else {
      return [];
    }
  }

  static propertyExists<T>(object: T, property: string): boolean {
    if(!object || typeof object !== 'object') return false
    return !!Object.getOwnPropertyDescriptor(object, property)
  }
}
