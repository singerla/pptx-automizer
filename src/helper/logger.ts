/**
 * Logging verbosity:
 * 0: errors only
 * 1: errors and warnings (default)
 * 2: everything, including info and debug output
 */
export type Verbosity = 0 | 1 | 2;

/**
 * Minimal logging interface used by pptx-automizer.
 * Inject a custom implementation via `AutomizerParams.logger`
 * to route library output into your own logging stack.
 */
export interface ILogger {
  error(message: unknown, ...details: unknown[]): void;
  warn(message: unknown, ...details: unknown[]): void;
  info(message: unknown, ...details: unknown[]): void;
  debug(message: unknown, ...details: unknown[]): void;
}

/**
 * Default logger. Writes to the console, filtered by verbosity.
 * Errors are always written; warnings at verbosity >= 1;
 * info and debug output at verbosity 2.
 */
export class ConsoleLogger implements ILogger {
  verbosity: Verbosity;

  constructor(verbosity: Verbosity = 1) {
    this.verbosity = verbosity;
  }

  error(message: unknown, ...details: unknown[]): void {
    console.error('[pptx-automizer]', message, ...details);
  }

  warn(message: unknown, ...details: unknown[]): void {
    if (this.verbosity >= 1) {
      console.warn('[pptx-automizer]', message, ...details);
    }
  }

  info(message: unknown, ...details: unknown[]): void {
    if (this.verbosity >= 2) {
      console.info('[pptx-automizer]', message, ...details);
    }
  }

  debug(message: unknown, ...details: unknown[]): void {
    if (this.verbosity >= 2) {
      console.debug('[pptx-automizer]', message, ...details);
    }
  }
}

/**
 * Discards all output. Inject via `AutomizerParams.logger` to keep
 * the library completely silent, e.g. when embedded in a server.
 */
export class NullLogger implements ILogger {
  error(): void {}
  warn(): void {}
  info(): void {}
  debug(): void {}
}

// The active logger is module-level state so that static helpers can log
// without holding an Automizer instance. It is replaced by the instance
// configured on Automizer. ROADMAP Phase 3 threads it through instead.
let activeLogger: ILogger = new ConsoleLogger();

export const setActiveLogger = (logger: ILogger): void => {
  activeLogger = logger;
};

/**
 * Logging facade for library internals: always delegates to the
 * logger currently configured on Automizer.
 */
export const log: ILogger = {
  error: (message, ...details) => activeLogger.error(message, ...details),
  warn: (message, ...details) => activeLogger.warn(message, ...details),
  info: (message, ...details) => activeLogger.info(message, ...details),
  debug: (message, ...details) => activeLogger.debug(message, ...details),
};
