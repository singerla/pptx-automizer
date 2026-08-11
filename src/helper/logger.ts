import { AsyncLocalStorage } from 'async_hooks';

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

// Static helpers log without holding an Automizer instance, and public
// modifier signatures cannot carry a logger. The instance configured on an
// Automizer is therefore scoped to its async call tree via AsyncLocalStorage:
// each Automizer wraps its entry points in runWithLogger(), so two instances
// running concurrently each log through their own logger.
const activeLogger = new AsyncLocalStorage<ILogger>();

// Fallback for calls outside any Automizer entry point.
const defaultLogger: ILogger = new ConsoleLogger();

/**
 * Runs `fn` with `logger` as the active logger for everything (a)waited
 * inside it. Used by Automizer to scope its configured logger to its own
 * pipeline; safe to nest.
 */
export const runWithLogger = <T>(logger: ILogger, fn: () => T): T => {
  return activeLogger.run(logger, fn);
};

/**
 * Logging facade for library internals: delegates to the logger of the
 * Automizer instance whose call tree we are in, or to a default
 * ConsoleLogger outside any instance context.
 */
export const log: ILogger = {
  error: (message, ...details) =>
    (activeLogger.getStore() ?? defaultLogger).error(message, ...details),
  warn: (message, ...details) =>
    (activeLogger.getStore() ?? defaultLogger).warn(message, ...details),
  info: (message, ...details) =>
    (activeLogger.getStore() ?? defaultLogger).info(message, ...details),
  debug: (message, ...details) =>
    (activeLogger.getStore() ?? defaultLogger).debug(message, ...details),
};
