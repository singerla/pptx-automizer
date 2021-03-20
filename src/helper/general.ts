export default class GeneralHelper {
  static arrayify<T>(input: T): T[] {
    if (input instanceof Array) {
      return input;
    } else if (input !== undefined) {
      return [input];
    } else {
      return [];
    }
  }
}
