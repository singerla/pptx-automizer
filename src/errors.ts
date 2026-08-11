/**
 * Base class for all errors thrown by pptx-automizer.
 * Allows consumers to catch library errors with `instanceof AutomizerError`.
 */
export class AutomizerError extends Error {
  constructor(message: string) {
    super(message);
    this.name = new.target.name;
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

/**
 * Thrown when a template file cannot be resolved in any of the
 * configured template directories.
 */
export class TemplateNotFoundError extends AutomizerError {
  readonly file: string;
  readonly searchedDirs: string[];

  constructor(file: string, searchedDirs: string[]) {
    super(
      `Template file not found: "${file}". Searched in: ${searchedDirs
        .map((dir) => (dir === '' ? '<working directory>' : dir))
        .join(', ')}`,
    );
    this.file = file;
    this.searchedDirs = searchedDirs;
  }
}

/**
 * Thrown when a slide cannot be resolved by number or creationId
 * in the given template.
 */
export class SlideNotFoundError extends AutomizerError {
  readonly slideIdentifier?: number | string;
  readonly templateName?: string;

  constructor(
    message: string,
    info?: { slideIdentifier?: number | string; templateName?: string },
  ) {
    super(message);
    this.slideIdentifier = info?.slideIdentifier;
    this.templateName = info?.templateName;
  }
}

/**
 * Thrown when a shape or xml element cannot be found, e.g. by an
 * element selector on a source slide.
 */
export class ElementNotFoundError extends AutomizerError {
  readonly selector?: string;
  readonly file?: string;

  constructor(message: string, info?: { selector?: string; file?: string }) {
    super(message);
    this.selector = info?.selector;
    this.file = info?.file;
  }
}

/**
 * Thrown on errors while reading from or writing to a pptx archive.
 */
export class ArchiveError extends AutomizerError {
  readonly file?: string;

  constructor(message: string, info?: { file?: string }) {
    super(message);
    this.file = info?.file;
  }
}

/**
 * Thrown when the output presentation cannot be produced as requested.
 */
export class OutputError extends AutomizerError {}

/**
 * Thrown when a user-provided modification callback throws.
 * The original exception is preserved on `cause`.
 * Set `AutomizerParams.continueOnError` to log and skip instead.
 */
export class CallbackError extends AutomizerError {
  readonly element?: string;
  readonly slideFile?: string;
  readonly cause: unknown;

  constructor(element: string, slideFile: string, cause: unknown) {
    const causeMessage =
      cause instanceof Error ? cause.message : String(cause);
    super(
      `Modification callback failed on element "${element}" (${slideFile}): ${causeMessage}. ` +
        `Set params.continueOnError to true to log and skip failing callbacks instead.`,
    );
    this.element = element;
    this.slideFile = slideFile;
    this.cause = cause;
  }
}
