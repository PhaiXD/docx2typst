/**
 * @file index.ts
 * Main entry point for the docx2typst library.
 * Exports public types, errors, utilities, and high-level conversion APIs.
 */

import { readDocxFile } from './reader.js';
import { extractText, resolveListItems } from './extractor/textExtractor.js';
import { extractNumberingMaps } from './extractor/numberingExtractor.js';
import { convertDocxToTypst } from './converter/typstConverter.js';
import type {
  DocxParserOptions,
  TypstConverterOptions,
  TypstDocument,
} from './types.js';

// Re-export all type definitions
export type {
  DocxParserOptions,
  RawXmlContent,
  TextRun,
  Paragraph,
  DocxDocument,
  TypstConverterOptions,
  TypstDocument,
  ListItemInfo,
  NumIdMap,
  AbstractNumMap,
} from './types.js';

// Re-export all custom error classes
export { DocxReadError, DocxParseError, type DocxErrorOptions } from './errors.js';

// Re-export parsing and extraction utilities
export { parseXml } from './utils/xmlParser.js';
export { readDocxFile } from './reader.js';
export { extractText, resolveListItems } from './extractor/textExtractor.js';
export { extractNumberingMaps } from './extractor/numberingExtractor.js';
export { convertDocxToTypst, getHeadingLevel } from './converter/typstConverter.js';

/**
 * Convenience function to load a .docx file from disk and convert its contents directly
 * to a plain text string.
 *
 * @param filePath - The path to the .docx file on the filesystem.
 * @param options - Optional parser options controlling style and whitespace extraction.
 * @returns A promise resolving to the plain text content of the document.
 * @throws {DocxReadError} If reading or opening the .docx archive fails.
 * @throws {DocxParseError} If the document XML content is malformed.
 *
 * @example
 * ```typescript
 * import { convertDocxToText } from 'docx2typst';
 *
 * const text = await convertDocxToText('./document.docx');
 * console.log(text);
 * ```
 */
export async function convertDocxToText(
  filePath: string,
  options?: DocxParserOptions,
): Promise<string> {
  const rawContent = await readDocxFile(filePath);
  let document = extractText(rawContent.documentXml, options);
  if (rawContent.numberingXml) {
    const { numIdMap, abstractNumMap } = extractNumberingMaps(rawContent.numberingXml);
    document = resolveListItems(document, numIdMap, abstractNumMap);
  }
  return document.text;
}

/**
 * Convenience function to load a .docx file from disk and convert its contents directly
 * to a Typst document containing markup content and conversion statistics.
 *
 * @param filePath - The path to the .docx file on the filesystem.
 * @param options - Optional Typst converter options controlling spacing, escaping, and headers.
 * @returns A promise resolving to the converted TypstDocument.
 * @throws {DocxReadError} If reading or opening the .docx archive fails.
 * @throws {DocxParseError} If the document XML content is malformed.
 *
 * @example
 * ```typescript
 * import { convertDocxFile } from 'docx2typst';
 *
 * const typstDoc = await convertDocxFile('./document.docx', { includeHeader: true });
 * console.log(typstDoc.content);
 * console.log(typstDoc.stats);
 * ```
 */
export async function convertDocxFile(
  filePath: string,
  options?: TypstConverterOptions,
): Promise<TypstDocument> {
  const rawContent = await readDocxFile(filePath);
  let document = extractText(rawContent.documentXml);
  if (rawContent.numberingXml) {
    const { numIdMap, abstractNumMap } = extractNumberingMaps(rawContent.numberingXml);
    document = resolveListItems(document, numIdMap, abstractNumMap);
  }
  return convertDocxToTypst(document, options);
}
