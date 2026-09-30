/**
 * @file reader.ts
 * Module for loading .docx archives from disk and extracting raw XML parts.
 */

import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { DocxReadError } from './errors.js';
import { parseXml } from './utils/xmlParser.js';
import type { DocxImage, RawXmlContent } from './types.js';

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
 * Type guard checking whether a value is a non-null Record object.
 *
 * @param value - Value to check.
 * @returns True if value is an object and not an array.
 */
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Normalizes a value that may be undefined, a single item, or an array into an array.
 *
 * @param value - Input value.
 * @returns An array of items.
 */
function toArray<T>(value: unknown): T[] {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? (value as T[]) : [value as T];
}

/**
 * Retrieves a property from a record matching one of several candidate keys.
 *
 * @param record - Target record.
 * @param keys - Candidate key names.
 * @returns The matching property value, or undefined.
 */
function getProperty(record: Record<string, unknown>, ...keys: string[]): unknown {
  for (const key of keys) {
    if (key in record && record[key] !== undefined) {
      return record[key];
    }
  }
  return undefined;
}

/**
 * Detects the MIME type of an image file based on its file extension.
 *
 * @param filePath - The path or filename of the image.
 * @returns The corresponding MIME type string, defaulting to 'image/png'.
 */
function detectMimeType(filePath: string): string {
  const ext = filePath.toLowerCase().split('.').pop() ?? '';
  switch (ext) {
    case 'png':
      return 'image/png';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'svg':
      return 'image/svg+xml';
    case 'emf':
      return 'image/emf';
    case 'wmf':
      return 'image/wmf';
    default:
      return 'image/png';
  }
}

/**
 * Parses OOXML relationships XML (`word/_rels/document.xml.rels`) to extract target mappings.
 *
 * @param relsXml - Raw XML string from the relationships part.
 * @returns Map of relationship ID to target path and relationship type.
 */
function parseRelationshipsForImages(
  relsXml: string,
): Map<string, { target: string; type: string }> {
  const map = new Map<string, { target: string; type: string }>();
  if (typeof relsXml !== 'string' || relsXml.trim().length === 0) {
    return map;
  }

  let parsed: Record<string, unknown>;
  try {
    parsed = parseXml(relsXml);
  } catch {
    return map;
  }

  const relsRoot = getProperty(parsed, 'Relationships', 'relationships', 'r:Relationships');
  const rootObj = isRecord(relsRoot) ? relsRoot : parsed;
  const relNodes = toArray<Record<string, unknown>>(
    getProperty(rootObj, 'Relationship', 'relationship', 'r:Relationship'),
  );

  for (const relNode of relNodes) {
    if (!isRecord(relNode)) {
      continue;
    }

    const id = getProperty(relNode, '@_Id', '@_id', '@Id', '@id', 'Id', 'id');
    const type = getProperty(relNode, '@_Type', '@_type', '@Type', '@type', 'Type', 'type');
    const target = getProperty(
      relNode,
      '@_Target',
      '@_target',
      '@Target',
      '@target',
      'Target',
      'target',
    );

    if (id !== undefined && type !== undefined && target !== undefined) {
      map.set(String(id), {
        target: String(target),
        type: String(type),
      });
    }
  }

  return map;
}

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

  // Parse relationships XML for image parts and construct imageMap
  const imageMap = new Map<string, DocxImage>();
  if (relationshipsXml) {
    const rels = parseRelationshipsForImages(relationshipsXml);
    for (const [relationshipId, { target, type }] of rels.entries()) {
      if (type.endsWith('/image')) {
        let zipPath = target.startsWith('word/') ? target : `word/${target.replace(/^\//, '')}`;
        if (zip.file(zipPath) === null && zip.file(target) !== null) {
          zipPath = target;
        }
        if (zip.file(zipPath) !== null) {
          const mimeType = detectMimeType(target);
          imageMap.set(relationshipId, {
            relationshipId,
            targetPath: target,
            zipPath,
            mimeType,
          });
        }
      }
    }
  }

  return {
    documentXml,
    ...(relationshipsXml !== undefined ? { relationshipsXml } : {}),
    ...(numberingXml !== undefined ? { numberingXml } : {}),
    imageMap,
    zipInstance: zip as unknown,
  };
}
