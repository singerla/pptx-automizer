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

export const last = <T>(arr: T[]): T => arr[arr.length - 1];

export interface Logger {
  verbosity: 0 | 1 | 2;
  target: 'console' | 'file';
  log: (
    message: string,
    verbosity: Logger['verbosity'],
    showStack?: boolean,
    target?: Logger['target'],
  ) => void;
}

export const Logger = <Logger>{
  verbosity: 1,
  target: 'console',
  log: (message, verbosity, showStack?, target?) => {
    if (verbosity > Logger.verbosity) {
      return;
    }
    target = target || Logger.target;
    if (target === 'console') {
      if (showStack) {
        vd(message);
      } else {
        console.log(message);
      }
    } else {
      // TODO: append message to a logfile
    }
  },
};

export const log = (message: string, verbosity: Logger['verbosity']) => {
  Logger.log(message, verbosity);
};

export const logDebug = (message: string, verbosity: Logger['verbosity']) => {
  Logger.log(message, verbosity, true);
};
