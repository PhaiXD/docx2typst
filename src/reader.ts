/**
 * @file reader.ts
 * Module for loading .docx archives from disk and extracting raw XML parts.
 */

import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { DocxReadError } from './errors.js';
import type { RawXmlContent } from './types.js';

/**
 * Path to the main document body XML inside the .docx ZIP package.
 */
const DOCUMENT_XML_PATH = 'word/document.xml';

/**
 * Path to the document relationships XML inside the .docx ZIP package.
 */
const DOCUMENT_RELS_XML_PATH = 'word/_rels/document.xml.rels';

/**
 * Path to the numbering definitions XML inside the .docx ZIP package.
 */
const NUMBERING_XML_PATH = 'word/numbering.xml';

/**
 * Reads a .docx file from the local filesystem and extracts its primary XML streams.
 *
 * @param filePath - The absolute or relative file path to the target .docx file.
 * @returns A promise resolving to the extracted raw XML content.
 * @throws {DocxReadError} If reading the file fails, if the file is not a valid ZIP archive,
 *                         or if the required `word/document.xml` part is missing.
 */
export async function readDocxFile(filePath: string): Promise<RawXmlContent> {
  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw new DocxReadError('A non-empty file path string must be provided.');
  }

  let fileBuffer: Buffer;
  try {
    fileBuffer = await readFile(filePath);
  } catch (error: unknown) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    throw new DocxReadError(`Failed to read file from disk at "${filePath}": ${errorMsg}`, {
      cause: error,
    });
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(fileBuffer);
  } catch (error: unknown) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    throw new DocxReadError(`File at "${filePath}" is not a valid ZIP/DOCX archive: ${errorMsg}`, {
      cause: error,
    });
  }

  const documentXmlEntry = zip.file(DOCUMENT_XML_PATH);
  if (!documentXmlEntry) {
    throw new DocxReadError(
      `Invalid .docx package at "${filePath}": missing required entry "${DOCUMENT_XML_PATH}".`,
    );
  }

  let documentXml: string;
  try {
    documentXml = await documentXmlEntry.async('string');
  } catch (error: unknown) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    throw new DocxReadError(
      `Failed to extract "${DOCUMENT_XML_PATH}" from "${filePath}": ${errorMsg}`,
      { cause: error },
    );
  }

  let relationshipsXml: string | undefined;
  const relsEntry = zip.file(DOCUMENT_RELS_XML_PATH);
  if (relsEntry) {
    try {
      relationshipsXml = await relsEntry.async('string');
    } catch {
      // Relationships are optional; proceed without relationships if read fails
      relationshipsXml = undefined;
    }
  }

  let numberingXml: string | undefined;
  const numberingEntry = zip.file(NUMBERING_XML_PATH);
  if (numberingEntry) {
    try {
      numberingXml = await numberingEntry.async('string');
    } catch {
      // Numbering is optional; proceed without numbering if read fails
      numberingXml = undefined;
    }
  }

  return {
    documentXml,
    ...(relationshipsXml !== undefined ? { relationshipsXml } : {}),
    ...(numberingXml !== undefined ? { numberingXml } : {}),
  };
}
