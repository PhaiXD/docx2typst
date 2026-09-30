/**
 * @file imageExtractor.ts
 * Module for extracting images and image metadata from OOXML DrawingML structures.
 */

import { basename } from 'node:path';
import { mkdir, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { DocxImage } from '../types.js';

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
 * Extracts relationship IDs from a drawing container node (`wp:inline` or `wp:anchor`).
 *
 * Navigates: container -> `a:graphic` -> `a:graphicData` -> `pic:pic` -> `pic:blipFill` -> `a:blip @r:embed`.
 *
 * @param container - The inline or anchor drawing container node.
 * @returns Array of relationship ID strings found in the container.
 */
function extractRelIdsFromContainer(container: Record<string, unknown>): string[] {
  const relIds: string[] = [];
  const graphicNodes = toArray<Record<string, unknown>>(
    getProperty(container, 'a:graphic', 'graphic'),
  );

  for (const graphic of graphicNodes) {
    if (!isRecord(graphic)) {
      continue;
    }
    const graphicDataNodes = toArray<Record<string, unknown>>(
      getProperty(graphic, 'a:graphicData', 'graphicData'),
    );

    for (const graphicData of graphicDataNodes) {
      if (!isRecord(graphicData)) {
        continue;
      }
      const picNodes = toArray<Record<string, unknown>>(
        getProperty(graphicData, 'pic:pic', 'pic'),
      );

      for (const pic of picNodes) {
        if (!isRecord(pic)) {
          continue;
        }
        const blipFillNodes = toArray<Record<string, unknown>>(
          getProperty(pic, 'pic:blipFill', 'blipFill'),
        );

        for (const blipFill of blipFillNodes) {
          if (!isRecord(blipFill)) {
            continue;
          }
          const blipNodes = toArray<Record<string, unknown>>(
            getProperty(blipFill, 'a:blip', 'blip'),
          );

          for (const blip of blipNodes) {
            if (!isRecord(blip)) {
              continue;
            }
            const rEmbed = getProperty(
              blip,
              '@_r:embed',
              '@_embed',
              '@r:embed',
              '@embed',
              'r:embed',
              'embed',
            );
            if (rEmbed !== undefined && rEmbed !== null) {
              const strVal = String(rEmbed).trim();
              if (strVal.length > 0) {
                relIds.push(strVal);
              }
            }
          }
        }
      }
    }
  }

  return relIds;
}

/**
 * Searches a run node (`w:r`) for embedded image drawing relationships.
 *
 * Looks for `w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:blip @r:embed`.
 * Also checks `w:drawing > wp:anchor` as an alternative to `wp:inline`.
 * Navigates properties using both namespace and non-namespace keys.
 *
 * @param rNode - Parsed run node representation.
 * @returns Array of relationship IDs found (may be empty).
 */
export function findImagesInRun(rNode: Record<string, unknown>): string[] {
  if (!isRecord(rNode)) {
    return [];
  }

  const drawingNodes = toArray<Record<string, unknown>>(
    getProperty(rNode, 'w:drawing', 'drawing'),
  );

  const relIds: string[] = [];

  for (const drawing of drawingNodes) {
    if (!isRecord(drawing)) {
      continue;
    }
    const inlines = toArray<Record<string, unknown>>(
      getProperty(drawing, 'wp:inline', 'inline'),
    );
    const anchors = toArray<Record<string, unknown>>(
      getProperty(drawing, 'wp:anchor', 'anchor'),
    );
    const containers = [...inlines, ...anchors];

    for (const container of containers) {
      if (!isRecord(container)) {
        continue;
      }
      relIds.push(...extractRelIdsFromContainer(container));
    }
  }

  return relIds;
}

/**
 * Extracts image metadata for images found within a run node.
 *
 * Calls `findImagesInRun` to discover relationship IDs, resolves each against `imageMap`,
 * and extracts dimension (`wp:extent @cx, @cy`) and alt text (`wp:docPr @descr` or `@title`)
 * metadata. Returns deep copies without mutating the `imageMap` entries.
 *
 * @param rNode - Parsed run node representation.
 * @param imageMap - Mapping from relationship ID to base DocxImage metadata.
 * @returns Array of DocxImage objects with resolved dimensions and alt text.
 */
export function extractImageMetadata(
  rNode: Record<string, unknown>,
  imageMap: Map<string, DocxImage>,
): DocxImage[] {
  if (!isRecord(rNode) || !imageMap || imageMap.size === 0) {
    return [];
  }

  const foundRelIds = findImagesInRun(rNode);
  if (foundRelIds.length === 0) {
    return [];
  }

  const drawingNodes = toArray<Record<string, unknown>>(
    getProperty(rNode, 'w:drawing', 'drawing'),
  );

  const result: DocxImage[] = [];

  for (const drawing of drawingNodes) {
    if (!isRecord(drawing)) {
      continue;
    }
    const inlines = toArray<Record<string, unknown>>(
      getProperty(drawing, 'wp:inline', 'inline'),
    );
    const anchors = toArray<Record<string, unknown>>(
      getProperty(drawing, 'wp:anchor', 'anchor'),
    );
    const containers = [...inlines, ...anchors];

    for (const container of containers) {
      if (!isRecord(container)) {
        continue;
      }

      const relIds = extractRelIdsFromContainer(container);
      if (relIds.length === 0) {
        continue;
      }

      // Extract width and height from wp:extent @cx and @cy
      const extentNode = getProperty(container, 'wp:extent', 'extent');
      let widthEmu: number | undefined;
      let heightEmu: number | undefined;

      if (isRecord(extentNode)) {
        const cx = getProperty(extentNode, '@_cx', '@cx', 'cx');
        if (cx !== undefined && cx !== null) {
          const parsedCx = parseInt(String(cx), 10);
          if (!isNaN(parsedCx)) {
            widthEmu = parsedCx;
          }
        }
        const cy = getProperty(extentNode, '@_cy', '@cy', 'cy');
        if (cy !== undefined && cy !== null) {
          const parsedCy = parseInt(String(cy), 10);
          if (!isNaN(parsedCy)) {
            heightEmu = parsedCy;
          }
        }
      }

      // Extract altText from wp:docPr @descr or @title
      const docPrNode = getProperty(container, 'wp:docPr', 'docPr');
      let altText: string | undefined;

      if (isRecord(docPrNode)) {
        const descr = getProperty(docPrNode, '@_descr', '@descr', 'descr');
        const title = getProperty(docPrNode, '@_title', '@title', 'title');

        if (descr !== undefined && descr !== null && String(descr).trim().length > 0) {
          altText = String(descr);
        } else if (title !== undefined && title !== null && String(title).trim().length > 0) {
          altText = String(title);
        }
      }

      for (const relId of relIds) {
        const baseImage = imageMap.get(relId);
        if (!baseImage) {
          continue;
        }

        // Deep copy the base DocxImage to avoid mutating imageMap entries
        const imageCopy: DocxImage = {
          ...baseImage,
          ...(widthEmu !== undefined ? { widthEmu } : {}),
          ...(heightEmu !== undefined ? { heightEmu } : {}),
          ...(altText !== undefined ? { altText } : {}),
        };
        result.push(imageCopy);
      }
    }
  }

  return result;
}

/**
 * Extracts image binary data from a DOCX zip archive and writes files to the filesystem.
 *
 * @param images - Array of DocxImage items to save.
 * @param zipInstance - The JSZip archive instance, typed as unknown and cast internally.
 * @param outputDir - Destination directory path for extracted images.
 * @returns A promise resolving to a Map of `zipPath` to saved file path.
 */
export async function saveImages(
  images: DocxImage[],
  zipInstance: unknown,
  outputDir: string,
): Promise<Map<string, string>> {
  const savedMap = new Map<string, string>();
  if (!images || images.length === 0 || !zipInstance) {
    return savedMap;
  }

  // Cast zipInstance from unknown to JSZip instance to access zip.file()
  const zip = zipInstance as JSZip;
  if (typeof zip.file !== 'function') {
    return savedMap;
  }

  if (outputDir.trim().length > 0) {
    await mkdir(outputDir, { recursive: true });
  }

  for (const image of images) {
    const fileEntry = zip.file(image.zipPath);
    if (!fileEntry) {
      continue;
    }

    const data = await fileEntry.async('uint8array');
    const filename = basename(image.zipPath.replace(/\\/g, '/'));
    const cleanOutputDir = outputDir.replace(/[/\\]+$/, '');
    const savedPath = cleanOutputDir.length > 0 ? `${cleanOutputDir}/${filename}` : filename;

    await writeFile(savedPath, data);
    savedMap.set(image.zipPath, savedPath);
  }

  return savedMap;
}
