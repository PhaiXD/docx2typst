/**
 * @file xmlParser.ts
 * XML parsing utility configuring fast-xml-parser for OOXML document parsing.
 */

import { XMLParser, XMLValidator } from 'fast-xml-parser';
import { DocxParseError } from '../errors.js';

/**
 * Shared instance of XMLParser configured for OOXML schemas.
 */
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: true,
  trimValues: true,
});

/**
 * Parses an XML string into a structured JavaScript object representation.
 *
 * @param xmlString - The raw XML content to parse.
 * @returns Parsed XML tree as a generic key-value object.
 * @throws {DocxParseError} If the XML string cannot be parsed or produces invalid output.
 */
export function parseXml(xmlString: string): Record<string, unknown> {
  if (typeof xmlString !== 'string' || xmlString.trim().length === 0) {
    throw new DocxParseError('Invalid XML input: expected a non-empty string.');
  }

  const validation = XMLValidator.validate(xmlString);
  if (validation !== true) {
    const errorMsg =
      typeof validation === 'object' && validation.err
        ? validation.err.msg
        : 'Invalid XML structure';
    throw new DocxParseError(`XML validation error: ${errorMsg}`);
  }

  try {
    const result: unknown = parser.parse(xmlString);

    if (result === null || typeof result !== 'object') {
      throw new DocxParseError('Failed to parse XML: output is not a valid object.');
    }

    return result as Record<string, unknown>;
  } catch (error: unknown) {
    if (error instanceof DocxParseError) {
      throw error;
    }
    throw new DocxParseError('Error parsing XML content', { cause: error });
  }
}

/**
 * Shared instance of XMLParser configured with preserveOrder: true for document-order body traversal.
 */
const orderedParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: true,
  trimValues: true,
});

/**
 * Parses an XML string into an ordered array of element nodes preserving sibling document order.
 *
 * @param xmlString - The raw XML content to parse.
 * @returns Parsed XML tree as an array of ordered element objects.
 * @throws {DocxParseError} If the XML string cannot be parsed or produces invalid output.
 */
export function parseXmlPreserveOrder(xmlString: string): Array<Record<string, unknown>> {
  if (typeof xmlString !== 'string' || xmlString.trim().length === 0) {
    throw new DocxParseError('Invalid XML input: expected a non-empty string.');
  }

  const validation = XMLValidator.validate(xmlString);
  if (validation !== true) {
    const errorMsg =
      typeof validation === 'object' && validation.err
        ? validation.err.msg
        : 'Invalid XML structure';
    throw new DocxParseError(`XML validation error: ${errorMsg}`);
  }

  try {
    const result: unknown = orderedParser.parse(xmlString);

    if (!Array.isArray(result)) {
      throw new DocxParseError('Failed to parse XML: output is not a valid array.');
    }

    return result as Array<Record<string, unknown>>;
  } catch (error: unknown) {
    if (error instanceof DocxParseError) {
      throw error;
    }
    throw new DocxParseError('Error parsing XML content', { cause: error });
  }
}

