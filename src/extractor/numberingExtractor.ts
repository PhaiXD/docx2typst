/**
 * @file numberingExtractor.ts
 * Module for extracting numbering definitions and mappings from WordprocessingML numbering XML.
 */

import { parseXml } from '../utils/xmlParser.js';
import type { AbstractNumFmtMap, AbstractNumMap, NumIdMap } from '../types.js';

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
      if (typeof val === 'string') {
        const parsed = parseInt(val, 10);
        return isNaN(parsed) ? undefined : parsed;
      }
      if (isRecord(val)) {
        const innerVal = getProperty(val, '@_w:val', '@_val', '@w:val', 'val', '#text');
        if (innerVal !== undefined && innerVal !== null) {
          if (typeof innerVal === 'number') return Math.trunc(innerVal);
          const parsed = parseInt(String(innerVal), 10);
          return isNaN(parsed) ? undefined : parsed;
        }
      }
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

  /**
   * Mapping from compound key `"${abstractNumId}:${level}"` to numbering format string (`w:numFmt @w:val`).
   */
  abstractNumFmtMap: AbstractNumFmtMap;
}

/**
 * Parses OOXML numbering XML (`word/numbering.xml`) and builds mapping tables for list resolution.
 *
 * Builds three maps:
 * 1. `abstractNumMap`: Maps `"${abstractNumId}:${ilvl}"` to `'bullet'` or `'ordered'` based on
 *    the level's numbering format (`w:numFmt @w:val`). 'bullet' maps to `'bullet'`, all other
 *    formats (e.g. 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman') map to `'ordered'`.
 * 2. `abstractNumFmtMap`: Maps `"${abstractNumId}:${ilvl}"` to raw numFmt string.
 * 3. `numIdMap`: Maps numbering instance ID (`w:num @w:numId`) to its corresponding abstract
 *    definition ID (`w:abstractNumId @w:val`).
 *
 * Gracefully handles empty, missing, or malformed XML by returning empty maps without throwing errors.
 *
 * @param numberingXml - Raw XML content from `word/numbering.xml`.
 * @returns An object containing `numIdMap`, `abstractNumMap`, and `abstractNumFmtMap`.
 */
export function extractNumberingMaps(numberingXml: string): ExtractedNumberingMaps {
  const numIdMap: NumIdMap = new Map();
  const abstractNumMap: AbstractNumMap = new Map();
  const abstractNumFmtMap: AbstractNumFmtMap = new Map();

  if (typeof numberingXml !== 'string' || numberingXml.trim().length === 0) {
    return { numIdMap, abstractNumMap, abstractNumFmtMap };
  }

  let parsed: Record<string, unknown>;
  try {
    parsed = parseXml(numberingXml);
  } catch {
    return { numIdMap, abstractNumMap, abstractNumFmtMap };
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

    let abstractNumId = parseIntAttribute(
      abstractNumNode,
      '@_w:abstractNumId',
      '@_abstractNumId',
      '@w:abstractNumId',
      'abstractNumId',
      '@_w:val',
      '@_val',
      '@w:val',
      'val',
    );

    if (abstractNumId === undefined) {
      for (const key of Object.keys(abstractNumNode)) {
        if (key.toLowerCase().includes('abstractnumid')) {
          const val = abstractNumNode[key];
          const parsed = parseIntAttribute(val);
          if (parsed !== undefined) {
            abstractNumId = parsed;
            break;
          }
        }
      }
    }

    if (abstractNumId === undefined) {
      console.warn('Warning: Could not extract abstractNumId from abstractNum node', abstractNumNode);
      continue;
    }

    // Process levels: <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
    const lvlRaw = getProperty(abstractNumNode, 'w:lvl', 'lvl');
    let lvlNodes: Record<string, unknown>[] = [];
    if (isRecord(lvlRaw) && !Array.isArray(lvlRaw)) {
      const keys = Object.keys(lvlRaw);
      const isIndexedByLvl = keys.length > 0 && keys.every((k) => /^\d+$/.test(k) || k === ':@');
      if (isIndexedByLvl) {
        for (const k of keys) {
          if (k !== ':@' && isRecord(lvlRaw[k])) {
            lvlNodes.push(lvlRaw[k] as Record<string, unknown>);
          }
        }
      } else {
        lvlNodes = [lvlRaw];
      }
    } else {
      lvlNodes = toArray<Record<string, unknown>>(lvlRaw);
    }

    for (let lvlIdx = 0; lvlIdx < lvlNodes.length; lvlIdx++) {
      const lvlNode = lvlNodes[lvlIdx];
      if (!isRecord(lvlNode)) {
        continue;
      }

      let ilvl = parseIntAttribute(
        lvlNode,
        '@_w:ilvl',
        '@_ilvl',
        '@w:ilvl',
        'ilvl',
        'w:ilvl',
        '@_w:val',
        '@_val',
        '@w:val',
        'val',
      );

      if (ilvl === undefined) {
        const ilvlChild = getProperty(lvlNode, 'w:ilvl', 'ilvl');
        if (ilvlChild !== undefined) {
          ilvl = parseIntAttribute(ilvlChild, '@_w:val', '@_val', '@w:val', 'val', 'val');
        }
      }

      if (ilvl === undefined) {
        for (const key of Object.keys(lvlNode)) {
          if (key.toLowerCase().includes('ilvl')) {
            const val = lvlNode[key];
            const parsed = parseIntAttribute(val, '@_w:val', '@_val', '@w:val', 'val');
            if (parsed !== undefined) {
              ilvl = parsed;
              break;
            }
          }
        }
      }

      if (ilvl === undefined) {
        ilvl = lvlIdx;
      }

      // Check numFmt element: <w:numFmt w:val="bullet"/>
      const numFmtNode = getProperty(lvlNode, 'w:numFmt', 'numFmt', '@_w:numFmt', '@_numFmt');
      let numFmtVal: unknown;

      if (isRecord(numFmtNode)) {
        numFmtVal = getProperty(numFmtNode, '@_w:val', '@_val', '@w:val', 'val', '#text');
        if (numFmtVal === undefined) {
          for (const key of Object.keys(numFmtNode)) {
            if (key.toLowerCase().includes('val')) {
              numFmtVal = numFmtNode[key];
              break;
            }
          }
        }
      } else if (typeof numFmtNode === 'string' || typeof numFmtNode === 'number') {
        numFmtVal = numFmtNode;
      }

      if (numFmtVal === undefined) {
        for (const key of Object.keys(lvlNode)) {
          if (key.toLowerCase().includes('numfmt')) {
            const candidate = lvlNode[key];
            if (isRecord(candidate)) {
              numFmtVal = getProperty(candidate, '@_w:val', '@_val', '@w:val', 'val', '#text');
            } else {
              numFmtVal = candidate;
            }
            if (numFmtVal !== undefined) break;
          }
        }
      }

      const numFmtStr =
        numFmtVal !== undefined && numFmtVal !== null
          ? String(numFmtVal).toLowerCase().trim()
          : '';

      const listType: 'bullet' | 'ordered' = numFmtStr === 'bullet' ? 'bullet' : 'ordered';
      abstractNumMap.set(`${abstractNumId}:${ilvl}`, listType);
      abstractNumFmtMap.set(
        `${abstractNumId}:${ilvl}`,
        numFmtStr || (listType === 'bullet' ? 'bullet' : 'decimal'),
      );
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

    let numId = parseIntAttribute(
      numNode,
      '@_w:numId',
      '@_numId',
      '@w:numId',
      'numId',
      '@_w:val',
      '@_val',
      '@w:val',
      'val',
    );

    if (numId === undefined) {
      for (const key of Object.keys(numNode)) {
        if (key.toLowerCase().includes('numid')) {
          const val = numNode[key];
          const parsed = parseIntAttribute(val);
          if (parsed !== undefined) {
            numId = parsed;
            break;
          }
        }
      }
    }

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
        '@w:abstractNumId',
        'abstractNumId',
      );
      if (abstractNumId === undefined) {
        for (const key of Object.keys(abstractNumIdChild)) {
          if (key.toLowerCase().includes('val') || key.toLowerCase().includes('abstractnumid')) {
            const val = abstractNumIdChild[key];
            const parsed = parseIntAttribute(val);
            if (parsed !== undefined) {
              abstractNumId = parsed;
              break;
            }
          }
        }
      }
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

    if (abstractNumId === undefined) {
      for (const key of Object.keys(numNode)) {
        if (key.toLowerCase().includes('abstractnumid')) {
          const val = numNode[key];
          const parsed = parseIntAttribute(val, '@_w:val', '@_val', '@w:val', 'val');
          if (parsed !== undefined) {
            abstractNumId = parsed;
            break;
          }
        }
      }
    }

    if (abstractNumId !== undefined) {
      numIdMap.set(numId, abstractNumId);
    }
  }

  return {
    numIdMap,
    abstractNumMap,
    abstractNumFmtMap,
  };
}
