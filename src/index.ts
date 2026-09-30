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
  ImageExtractionOptions,
  DocxImage,
  RawXmlContent,
  TextRun,
  Paragraph,
  DocxDocument,
  TypstConverterOptions,
  TypstDocument,
  ListItemInfo,
  NumIdMap,
  AbstractNumMap,
  TableCell,
  TableRow,
  DocxTable,
} from './types.js';

// Re-export all custom error classes
export { DocxReadError, DocxParseError, type DocxErrorOptions } from './errors.js';

// Re-export parsing and extraction utilities
export { parseXml } from './utils/xmlParser.js';
export { readDocxFile } from './reader.js';
export { extractText, resolveListItems } from './extractor/textExtractor.js';
export { extractNumberingMaps } from './extractor/numberingExtractor.js';
export { extractTables } from './extractor/tableExtractor.js';
export {
  findImagesInRun,
  extractImageMetadata,
  saveImages,
} from './extractor/imageExtractor.js';
export {
  convertDocxToTypst,
  convertTableToTypst,
  getHeadingLevel,
} from './converter/typstConverter.js';

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
 * @example Basic plain text extraction:
 * ```typescript
 * import { convertDocxToText } from 'docx2typst';
 *
 * const text = await convertDocxToText('./notes.docx');
 * console.log('Document text:\n', text);
 * ```
 *
 * @example Error handling and whitespace options:
 * ```typescript
 * import { convertDocxToText, DocxReadError, DocxParseError } from 'docx2typst';
 *
 * try {
 *   const plainText = await convertDocxToText('./report.docx', {
 *     preserveWhitespace: true,
 *     includeStyles: false,
 *   });
 *   console.log(plainText);
 * } catch (error) {
 *   if (error instanceof DocxReadError) {
 *     console.error('Failed to read DOCX package:', error.message);
 *   } else if (error instanceof DocxParseError) {
 *     console.error('Invalid document XML structure:', error.message);
 *   }
 * }
 * ```
 */
export async function convertDocxToText(
  filePath: string,
  options?: DocxParserOptions,
): Promise<string> {
  const rawContent = await readDocxFile(filePath);
  let document = extractText(rawContent.documentXml, options, rawContent.imageMap);
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
 * @example Converting a DOCX file and writing Typst output to disk:
 * ```typescript
 * import { writeFile } from 'node:fs/promises';
 * import { convertDocxFile } from 'docx2typst';
 *
 * const typstDoc = await convertDocxFile('./document.docx');
 * await writeFile('./document.typ', typstDoc.content, 'utf-8');
 * console.log(`Converted ${typstDoc.stats.paragraphCount} paragraphs and ${typstDoc.stats.headingCount} headings.`);
 * ```
 *
 * @example Custom configuration with image directory and header comments:
 * ```typescript
 * import { convertDocxFile } from 'docx2typst';
 *
 * const typstDoc = await convertDocxFile('./document.docx', {
 *   includeHeader: true,
 *   imageOutputDir: 'assets/images',
 *   escapeSpecialChars: true,
 *   paragraphSpacing: true,
 * });
 * console.log(typstDoc.content);
 * console.log('Conversion statistics:', typstDoc.stats);
 * ```
 */
export async function convertDocxFile(
  filePath: string,
  options?: TypstConverterOptions,
): Promise<TypstDocument> {
  const rawContent = await readDocxFile(filePath);
  let document = extractText(rawContent.documentXml, undefined, rawContent.imageMap);
  if (rawContent.numberingXml) {
    const { numIdMap, abstractNumMap } = extractNumberingMaps(rawContent.numberingXml);
    document = resolveListItems(document, numIdMap, abstractNumMap);
  }
  return convertDocxToTypst(document, options);
}

