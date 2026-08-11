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
