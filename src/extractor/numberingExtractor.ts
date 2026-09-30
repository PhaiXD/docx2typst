/**
 * @file numberingExtractor.ts
 * Module for extracting numbering definitions and mappings from WordprocessingML numbering XML.
 */

import { parseXml } from '../utils/xmlParser.js';
import type { AbstractNumMap, NumIdMap } from '../types.js';

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
 * Extracts integer value from an attribute or node representation.
 *
 * @param node - Node or attribute value.
 * @param attrKeys - Candidate attribute keys if node is a record.
 * @returns Parsed integer or undefined if not extractable.
 */
function parseIntAttribute(node: unknown, ...attrKeys: string[]): number | undefined {
  if (node === undefined || node === null) {
    return undefined;
  }
  if (typeof node === 'number') {
    return Math.trunc(node);
  }
  if (typeof node === 'string') {
    const parsed = parseInt(node, 10);
    return isNaN(parsed) ? undefined : parsed;
  }
  if (isRecord(node)) {
    const val = getProperty(node, ...attrKeys);
    if (val !== undefined && val !== null) {
      if (typeof val === 'number') {
        return Math.trunc(val);
      }
      const parsed = parseInt(String(val), 10);
      return isNaN(parsed) ? undefined : parsed;
    }
  }
  return undefined;
}

/**
 * Result returned by {@link extractNumberingMaps}.
 */
export interface ExtractedNumberingMaps {
  /**
   * Mapping from numbering instance ID (`numId`) to abstract numbering definition ID (`abstractNumId`).
   */
  numIdMap: NumIdMap;

  /**
   * Mapping from compound key `"${abstractNumId}:${level}"` to list type (`bullet` or `ordered`).
   */
  abstractNumMap: AbstractNumMap;
}

/**
 * Parses OOXML numbering XML (`word/numbering.xml`) and builds mapping tables for list resolution.
 *
 * Builds two maps:
 * 1. `abstractNumMap`: Maps `"${abstractNumId}:${ilvl}"` to `'bullet'` or `'ordered'` based on
 *    the level's numbering format (`w:numFmt @w:val`). 'bullet' maps to `'bullet'`, all other
 *    formats (e.g. 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman') map to `'ordered'`.
 * 2. `numIdMap`: Maps numbering instance ID (`w:num @w:numId`) to its corresponding abstract
 *    definition ID (`w:abstractNumId @w:val`).
 *
 * Gracefully handles empty, missing, or malformed XML by returning empty maps without throwing errors.
 *
 * @param numberingXml - Raw XML content from `word/numbering.xml`.
 * @returns An object containing `numIdMap` and `abstractNumMap`.
 */
export function extractNumberingMaps(numberingXml: string): ExtractedNumberingMaps {
  const numIdMap: NumIdMap = new Map();
  const abstractNumMap: AbstractNumMap = new Map();

  if (typeof numberingXml !== 'string' || numberingXml.trim().length === 0) {
    return { numIdMap, abstractNumMap };
  }

  let parsed: Record<string, unknown>;
  try {
    parsed = parseXml(numberingXml);
  } catch {
    return { numIdMap, abstractNumMap };
  }

  const numberingRoot = getProperty(parsed, 'w:numbering', 'numbering');
  const rootObj = isRecord(numberingRoot) ? numberingRoot : parsed;

  // 1. Process abstractNum elements: <w:abstractNum w:abstractNumId="0">
  const abstractNumNodes = toArray<Record<string, unknown>>(
    getProperty(rootObj, 'w:abstractNum', 'abstractNum'),
  );

  for (const abstractNumNode of abstractNumNodes) {
    if (!isRecord(abstractNumNode)) {
      continue;
    }

    const abstractNumId = parseIntAttribute(
      abstractNumNode,
      '@_w:abstractNumId',
      '@_abstractNumId',
      '@w:abstractNumId',
      'abstractNumId',
      '@_w:val',
      '@_val',
    );

    if (abstractNumId === undefined) {
      continue;
    }

    // Process levels: <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
    const lvlNodes = toArray<Record<string, unknown>>(
      getProperty(abstractNumNode, 'w:lvl', 'lvl'),
    );

    for (const lvlNode of lvlNodes) {
      if (!isRecord(lvlNode)) {
        continue;
      }

      const ilvl = parseIntAttribute(
        lvlNode,
        '@_w:ilvl',
        '@_ilvl',
        '@w:ilvl',
        'ilvl',
        '@_w:val',
        '@_val',
      );

      if (ilvl === undefined) {
        continue;
      }

      // Check numFmt element: <w:numFmt w:val="bullet"/>
      const numFmtNode = getProperty(lvlNode, 'w:numFmt', 'numFmt');
      let numFmtVal: unknown;

      if (isRecord(numFmtNode)) {
        numFmtVal = getProperty(numFmtNode, '@_w:val', '@_val', '@w:val', 'val');
      } else if (typeof numFmtNode === 'string' || typeof numFmtNode === 'number') {
        numFmtVal = numFmtNode;
      }

      const numFmtStr =
        numFmtVal !== undefined && numFmtVal !== null
          ? String(numFmtVal).toLowerCase().trim()
          : '';

      const listType: 'bullet' | 'ordered' = numFmtStr === 'bullet' ? 'bullet' : 'ordered';
      abstractNumMap.set(`${abstractNumId}:${ilvl}`, listType);
    }
  }

  // 2. Process num elements: <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  const numNodes = toArray<Record<string, unknown>>(
    getProperty(rootObj, 'w:num', 'num'),
  );

  for (const numNode of numNodes) {
    if (!isRecord(numNode)) {
      continue;
    }

    const numId = parseIntAttribute(
      numNode,
      '@_w:numId',
      '@_numId',
      '@w:numId',
      'numId',
      '@_w:val',
      '@_val',
    );

    if (numId === undefined) {
      continue;
    }

    const abstractNumIdChild = getProperty(numNode, 'w:abstractNumId', 'abstractNumId');
    let abstractNumId: number | undefined;

    if (isRecord(abstractNumIdChild)) {
      abstractNumId = parseIntAttribute(
        abstractNumIdChild,
        '@_w:val',
        '@_val',
        '@w:val',
        'val',
        '@_w:abstractNumId',
        '@_abstractNumId',
      );
    } else if (abstractNumIdChild !== undefined) {
      abstractNumId = parseIntAttribute(abstractNumIdChild);
    } else {
      abstractNumId = parseIntAttribute(
        numNode,
        '@_w:abstractNumId',
        '@_abstractNumId',
        '@w:abstractNumId',
        'abstractNumId',
      );
    }

    if (abstractNumId !== undefined) {
      numIdMap.set(numId, abstractNumId);
    }
  }

  return {
    numIdMap,
    abstractNumMap,
  };
}
