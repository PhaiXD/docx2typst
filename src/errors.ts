/**
 * @file errors.ts
 * Custom error classes for DOCX reading and parsing operations.
 */

/**
 * Options for instantiating custom DOCX errors, supporting error chaining.
 */
export interface DocxErrorOptions {
  /**
   * The underlying cause of the error.
   */
  cause?: unknown;
}

/**
 * Error thrown when reading, opening, or extracting a .docx file fails.
 */
export class DocxReadError extends Error {
  /**
   * The root cause of the read failure, if available.
   */
  public readonly cause?: unknown;

  /**
   * Creates a new instance of DocxReadError.
   *
   * @param message - Human-readable error description.
   * @param options - Optional configuration including the underlying error cause.
   */
  constructor(message: string, options?: DocxErrorOptions) {
    super(message);
    this.name = 'DocxReadError';
    this.cause = options?.cause;

    // Restore prototype chain for instanceof checks across transpilations
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

/**
 * Error thrown when parsing OOXML markup or extracting document structures fails.
 */
export class DocxParseError extends Error {
  /**
   * The root cause of the parse failure, if available.
   */
  public readonly cause?: unknown;

  /**
   * Creates a new instance of DocxParseError.
   *
   * @param message - Human-readable error description.
   * @param options - Optional configuration including the underlying error cause.
   */
  constructor(message: string, options?: DocxErrorOptions) {
    super(message);
    this.name = 'DocxParseError';
    this.cause = options?.cause;

    // Restore prototype chain for instanceof checks across transpilations
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
