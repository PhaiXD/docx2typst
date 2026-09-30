/**
 * @file textExtractor.ts
 * Module for parsing OOXML WordprocessingML content and extracting structured text.
 */

import { parseXml } from '../utils/xmlParser.js';
import { extractTables } from './tableExtractor.js';
import { extractImageMetadata } from './imageExtractor.js';
import type {
  AbstractNumMap,
  DocxDocument,
  DocxImage,
  DocxParserOptions,
  ListItemInfo,
  NumIdMap,
  Paragraph,
  TextRun,
} from '../types.js';

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
 * Checks whether an OOXML toggle property (such as w:b or w:i) represents an active state.
 * In OOXML, `<w:b/>`, `<w:b w:val="true"/>`, `<w:b w:val="1"/>`, and `<w:b w:val="on"/>` are active.
 *
 * @param prop - The property node from parsed XML.
 * @returns True if active, false otherwise.
 */
function isToggleActive(prop: unknown): boolean {
  if (prop === undefined || prop === null) {
    return false;
  }
  if (typeof prop === 'boolean') {
    return prop;
  }
  if (typeof prop === 'number') {
    return prop !== 0;
  }
  if (typeof prop === 'string') {
    const s = prop.toLowerCase().trim();
    return s !== 'false' && s !== '0' && s !== 'off';
  }
  if (isRecord(prop)) {
    const val = getProperty(prop, '@_w:val', '@_val');
    if (val === undefined) {
      // Element present without attributes, e.g. <w:b/> -> defaults to true
      return true;
    }
    if (typeof val === 'boolean') {
      return val;
    }
    const valStr = String(val).toLowerCase().trim();
    return valStr !== 'false' && valStr !== '0' && valStr !== 'off' && valStr !== 'none';
  }
  return false;
}

/**
 * Checks whether an underline property (w:u) represents active underlining.
 *
 * @param prop - The property node from parsed XML.
 * @returns True if active, false otherwise.
 */
function isUnderlineActive(prop: unknown): boolean {
  if (prop === undefined || prop === null) {
    return false;
  }
  if (typeof prop === 'boolean') {
    return prop;
  }
  if (typeof prop === 'string') {
    const s = prop.toLowerCase().trim();
    return s !== 'none' && s !== 'false' && s !== '0';
  }
  if (isRecord(prop)) {
    const val = getProperty(prop, '@_w:val', '@_val');
    if (val === undefined) {
      return true;
    }
    const valStr = String(val).toLowerCase().trim();
    return valStr !== 'none' && valStr !== 'false' && valStr !== '0';
  }
  return false;
}

/**
 * Extracts raw string content from an OOXML text node (w:t), taking into account
 * attributes such as xml:space="preserve".
 *
 * @param tNode - The w:t node in parsed XML.
 * @returns The extracted string value.
 */
function extractTextFromTNode(tNode: unknown): string {
  if (tNode === undefined || tNode === null) {
    return '';
  }
  if (typeof tNode === 'string') {
    return tNode;
  }
  if (typeof tNode === 'number' || typeof tNode === 'boolean') {
    return String(tNode);
  }
  if (isRecord(tNode)) {
    const textVal = tNode['#text'];
    if (textVal !== undefined && textVal !== null) {
      return String(textVal);
    }
    return '';
  }
  return '';
}

/**
 * Extracts a single TextRun from a parsed w:r node.
 * If images are present in w:drawing and imageMap is provided, extracts DocxImage metadata.
 *
 * @param rNode - Parsed w:r node.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param paragraphImages - Optional array accumulating images in the paragraph.
 * @returns Extracted TextRun, or null if no valid run could be parsed.
 */
function extractRun(
  rNode: unknown,
  imageMap?: Map<string, DocxImage>,
  paragraphImages?: DocxImage[],
): TextRun | null {
  if (!isRecord(rNode)) {
    return null;
  }

  // Extract run properties (w:rPr / rPr)
  const rPrRaw = getProperty(rNode, 'w:rPr', 'rPr');
  const rPr = isRecord(rPrRaw) ? rPrRaw : undefined;

  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underline: boolean | undefined;

  if (rPr) {
    const bNode = getProperty(rPr, 'w:b', 'b');
    if (isToggleActive(bNode)) {
      bold = true;
    }

    const iNode = getProperty(rPr, 'w:i', 'i');
    if (isToggleActive(iNode)) {
      italic = true;
    }

    const uNode = getProperty(rPr, 'w:u', 'u');
    if (isUnderlineActive(uNode)) {
      underline = true;
    }
  }

  // Extract text nodes (w:t)
  const tNodes = toArray(getProperty(rNode, 'w:t', 't'));
  let text = '';
  for (const tNode of tNodes) {
    text += extractTextFromTNode(tNode);
  }

  // Handle special inline elements such as w:tab and w:br
  if (getProperty(rNode, 'w:tab', 'tab') !== undefined) {
    text += '\t';
  }
  if (getProperty(rNode, 'w:br', 'br') !== undefined) {
    text += '\n';
  }

  // Check for w:drawing in the run node
  const drawingNode = getProperty(rNode, 'w:drawing', 'drawing');
  if (drawingNode !== undefined && imageMap && paragraphImages && isRecord(rNode)) {
    const images = extractImageMetadata(rNode, imageMap);
    if (images.length > 0) {
      paragraphImages.push(...images);
    }
  }

  // Return run even if text is empty, as long as it was a valid run element
  const run: TextRun = {
    text,
    ...(bold ? { bold: true } : {}),
    ...(italic ? { italic: true } : {}),
    ...(underline ? { underline: true } : {}),
  };

  return run;
}

/**
 * Recursively extracts all run nodes (w:r) from a paragraph or container,
 * including runs nested inside hyperlinks (w:hyperlink).
 *
 * @param container - Parsed paragraph or child container node.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param paragraphImages - Optional array accumulating images in the paragraph.
 * @returns Array of parsed TextRun objects.
 */
function extractRunsFromContainer(
  container: Record<string, unknown>,
  imageMap?: Map<string, DocxImage>,
  paragraphImages?: DocxImage[],
): TextRun[] {
  const runs: TextRun[] = [];

  // Direct runs (w:r)
  const directRuns = toArray(getProperty(container, 'w:r', 'r'));
  for (const rNode of directRuns) {
    const run = extractRun(rNode, imageMap, paragraphImages);
    if (run !== null) {
      runs.push(run);
    }
  }

  // Runs nested inside hyperlinks (w:hyperlink)
  const hyperlinks = toArray<Record<string, unknown>>(
    getProperty(container, 'w:hyperlink', 'hyperlink'),
  );
  for (const hyperlink of hyperlinks) {
    if (isRecord(hyperlink)) {
      runs.push(...extractRunsFromContainer(hyperlink, imageMap, paragraphImages));
    }
  }

  return runs;
}

/**
 * Extracts a Paragraph object from a parsed w:p node.
 *
 * @param pNode - Parsed w:p XML node.
 * @param options - Parser options.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @returns Extracted Paragraph.
 */
export function extractParagraph(
  pNode: unknown,
  options?: DocxParserOptions,
  imageMap?: Map<string, DocxImage>,
): Paragraph {
  if (!isRecord(pNode)) {
    return {
      text: '',
      runs: [],
      isEmpty: true,
    };
  }

  // Extract paragraph properties (w:pPr / pPr)
  const pPrRaw = getProperty(pNode, 'w:pPr', 'pPr');
  const pPr = isRecord(pPrRaw) ? pPrRaw : undefined;

  let style: string | undefined;
  if (options?.includeStyles !== false && pPr) {
    const pStyleRaw = getProperty(pPr, 'w:pStyle', 'pStyle');
    if (isRecord(pStyleRaw)) {
      const styleVal = getProperty(pStyleRaw, '@_w:val', '@_val');
      if (styleVal !== undefined && styleVal !== null) {
        style = String(styleVal);
      }
    }
  }

  // Check for list item numbering (w:numPr / numPr)
  let listItem: ListItemInfo | undefined;
  if (pPr) {
    const numPrRaw = getProperty(pPr, 'w:numPr', 'numPr');
    if (isRecord(numPrRaw)) {
      let numId = 0;
      const numIdNode = getProperty(numPrRaw, 'w:numId', 'numId');
      if (isRecord(numIdNode)) {
        const val = getProperty(numIdNode, '@_w:val', '@_val', '@w:val', 'val');
        if (val !== undefined && val !== null) {
          const parsed = parseInt(String(val), 10);
          if (!isNaN(parsed)) {
            numId = parsed;
          }
        }
      } else if (typeof numIdNode === 'number') {
        numId = Math.trunc(numIdNode);
      } else if (typeof numIdNode === 'string') {
        const parsed = parseInt(numIdNode, 10);
        if (!isNaN(parsed)) {
          numId = parsed;
        }
      }

      // If numId is not 0, it is an active list item
      if (numId !== 0) {
        let level = 0;
        const ilvlNode = getProperty(numPrRaw, 'w:ilvl', 'ilvl');
        if (isRecord(ilvlNode)) {
          const val = getProperty(ilvlNode, '@_w:val', '@_val', '@w:val', 'val');
          if (val !== undefined && val !== null) {
            const parsed = parseInt(String(val), 10);
            if (!isNaN(parsed)) {
              level = parsed;
            }
          }
        } else if (typeof ilvlNode === 'number') {
          level = Math.trunc(ilvlNode);
        } else if (typeof ilvlNode === 'string') {
          const parsed = parseInt(ilvlNode, 10);
          if (!isNaN(parsed)) {
            level = parsed;
          }
        }

        listItem = {
          numId,
          level,
          abstractNumId: 0,
          listType: 'bullet',
        };
      }
    }
  }

  const paragraphImages: DocxImage[] = [];
  const runs = extractRunsFromContainer(pNode, imageMap, paragraphImages);
  const text = runs.map((r) => r.text).join('');
  const isEmpty = (runs.length === 0 || text.length === 0) && paragraphImages.length === 0;

  return {
    text,
    runs,
    ...(style !== undefined ? { style } : {}),
    ...(isEmpty ? { isEmpty: true } : {}),
    ...(listItem !== undefined ? { listItem } : {}),
    ...(paragraphImages.length > 0 ? { images: paragraphImages } : {}),
  };
}

/**
 * Parses raw WordprocessingML XML content into a structured DocxDocument.
 *
 * Navigates the OOXML hierarchy: `w:document` -> `w:body` -> `w:p` -> `w:r` -> `w:t`.
 * Handles formatting indicators (`w:b`, `w:i`, `w:u`), paragraph styles (`w:pStyle`),
 * empty paragraphs, whitespace preservation, and images.
 *
 * @param rawXml - Raw XML string from `word/document.xml`.
 * @param options - Optional parser configuration.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @returns The structured DocxDocument containing paragraphs and plain text.
 * @throws {DocxParseError} If the XML is malformed or cannot be parsed.
 */
export function extractText(
  rawXml: string,
  options?: DocxParserOptions,
  imageMap?: Map<string, DocxImage>,
): DocxDocument {
  const parsed = parseXml(rawXml);

  // Navigate to root document element
  const docRoot = getProperty(parsed, 'w:document', 'document');
  const rootObj = isRecord(docRoot) ? docRoot : parsed;

  // Navigate to body element
  const bodyRoot = getProperty(rootObj, 'w:body', 'body');
  const bodyObj = isRecord(bodyRoot) ? bodyRoot : rootObj;

  // Extract paragraphs (w:p)
  const pNodes = toArray(getProperty(bodyObj, 'w:p', 'p'));
  const paragraphs: Paragraph[] = pNodes.map((pNode) =>
    extractParagraph(pNode, options, imageMap),
  );

  const text = paragraphs.map((p) => p.text).join('\n');

  // Extract tables (w:tbl)
  const tables = isRecord(bodyObj) ? extractTables(bodyObj) : [];

  return {
    paragraphs,
    text,
    tables,
  };
}

/**
 * Resolves list item properties (abstractNumId and listType) across all paragraphs
 * in a DocxDocument using numbering definition maps.
 *
 * This function creates a new DocxDocument without mutating the input document.
 * For each paragraph with a listItem placeholder, it looks up abstractNumId from numIdMap,
 * and listType from abstractNumMap using key "${abstractNumId}:${level}".
 * If no matching list type is found, it defaults to 'bullet'.
 *
 * @param doc - The parsed DOCX document structure.
 * @param numIdMap - Mapping from numId to abstractNumId.
 * @param abstractNumMap - Mapping from "${abstractNumId}:${level}" to 'bullet' | 'ordered'.
 * @returns A new DocxDocument with resolved list items.
 */
export function resolveListItems(
  doc: DocxDocument,
  numIdMap: NumIdMap,
  abstractNumMap: AbstractNumMap,
): DocxDocument {
  const paragraphs: Paragraph[] = doc.paragraphs.map((para) => {
    if (!para.listItem) {
      return { ...para };
    }

    const { numId, level } = para.listItem;
    const abstractNumId = numIdMap.get(numId) ?? 0;
    const listType = abstractNumMap.get(`${abstractNumId}:${level}`) ?? 'bullet';

    const listItem: ListItemInfo = {
      abstractNumId,
      numId,
      level,
      listType,
    };

    return {
      ...para,
      listItem,
    };
  });

  return {
    ...doc,
    paragraphs,
  };
}
