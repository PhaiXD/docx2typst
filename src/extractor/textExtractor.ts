/**
 * @file textExtractor.ts
 * Module for parsing OOXML WordprocessingML content and extracting structured text.
 */

import { parseXml, parseXmlPreserveOrder } from '../utils/xmlParser.js';
import { extractTables } from './tableExtractor.js';
import { extractImageMetadata } from './imageExtractor.js';
import type {
  AbstractNumFmtMap,
  AbstractNumMap,
  BodyItem,
  DocxDocument,
  DocxImage,
  DocxParserOptions,
  DocxStyleInfo,
  DocxTable,
  ListItemInfo,
  NumIdMap,
  Paragraph,
  ParagraphIndent,
  SectionProperties,
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
 * Recursively discovers all w:txbxContent nodes within a drawing or shape container.
 *
 * @param node - Container node.
 * @returns Array of w:txbxContent record nodes found.
 */
function findTxbxContentNodes(node: unknown): Record<string, unknown>[] {
  const results: Record<string, unknown>[] = [];
  function search(curr: unknown): void {
    if (!isRecord(curr)) return;
    for (const [key, val] of Object.entries(curr)) {
      if (key === 'w:txbxContent' || key === 'txbxContent') {
        if (isRecord(val)) {
          results.push(val);
        } else if (Array.isArray(val)) {
          for (const item of val) {
            if (isRecord(item)) results.push(item);
          }
        }
      } else if (isRecord(val)) {
        search(val);
      } else if (Array.isArray(val)) {
        for (const item of val) {
          search(item);
        }
      }
    }
  }
  search(node);
  return results;
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
 * Retrieves a numeric attribute from a record matching one of several candidate keys.
 * Tries each key, parses as integer, and returns the first valid number or undefined.
 *
 * @param record - Target record.
 * @param keys - Candidate attribute keys.
 * @returns The parsed integer value, or undefined if none found or not a number.
 */
function getNumericAttr(record: Record<string, unknown>, ...keys: string[]): number | undefined {
  for (const key of keys) {
    if (key in record && record[key] !== undefined && record[key] !== null) {
      const val = parseInt(String(record[key]), 10);
      if (!isNaN(val)) {
        return val;
      }
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
  defaultColor?: string,
  hyperlinkMap?: Map<string, string>,
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
  let color: string | undefined;
  let highlight: string | undefined;
  let linkTarget: string | undefined;

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

    const colorNode = getProperty(rPr, 'w:color', 'color');
    if (isRecord(colorNode)) {
      const colorValRaw = getProperty(colorNode, '@_w:val', '@_val', '@w:val', 'val');
      if (colorValRaw !== undefined && colorValRaw !== 'auto') {
        const colorVal = String(colorValRaw).padStart(6, '0');
        if (colorVal !== '000000') {
          color = colorVal;
        }
      }
    } else if (defaultColor) {
      color = defaultColor;
    }

    const highlightNode = getProperty(rPr, 'w:highlight', 'highlight');
    if (isRecord(highlightNode)) {
      const hlVal = getProperty(highlightNode, '@_w:val', '@_val', '@w:val', 'val');
      if (hlVal !== undefined && hlVal !== 'none') {
        const ooxmlHighlightMap: Record<string, string> = {
          black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00',
          magenta: 'FF00FF', red: 'FF0000', yellow: 'FFFF00', white: 'FFFFFF',
          darkBlue: '000080', darkCyan: '008080', darkGreen: '008000',
          darkMagenta: '800080', darkRed: '800000', darkYellow: '808000',
          darkGray: '808080', lightGray: 'C0C0C0'
        };
        const colorName = String(hlVal);
        highlight = ooxmlHighlightMap[colorName] || colorName;
      }
    }

    const shdNode = getProperty(rPr, 'w:shd', 'shd');
    if (!highlight && isRecord(shdNode)) {
      const shdFill = getProperty(shdNode, '@_w:fill', '@_fill', '@w:fill', 'fill');
      if (shdFill !== undefined && shdFill !== 'auto') {
        highlight = String(shdFill).padStart(6, '0');
      }
    }

    const linkNode = getProperty(rPr, 'w:linkTarget', 'linkTarget');
    if (isRecord(linkNode)) {
      const linkId = String(getProperty(linkNode, '@_w:val', '@_val', '@w:val', 'val'));
      if (linkId && hyperlinkMap?.has(linkId)) {
        linkTarget = hyperlinkMap.get(linkId);
      }
    }
  } else if (defaultColor) {
    color = defaultColor;
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

  let pageBreak: boolean | undefined;
  const brNodes = toArray(getProperty(rNode, 'w:br', 'br'));
  for (const br of brNodes) {
    if (isRecord(br)) {
      const brType = getProperty(br, '@_w:type', '@_type', '@w:type', 'type');
      if (brType === 'page') {
        pageBreak = true;
      } else {
        text += '\n';
      }
    } else if (br !== undefined && br !== null) {
      text += '\n';
    }
  }

  // Check for w:drawing in the run node
  const drawingNode = getProperty(rNode, 'w:drawing', 'drawing');
  if (drawingNode !== undefined && imageMap && paragraphImages && isRecord(rNode)) {
    const images = extractImageMetadata(rNode, imageMap);
    if (images.length > 0) {
      paragraphImages.push(...images);
    }
  }

  // Check for horizontal line in w:pict
  let horizontalLine = false;
  const pictNode = getProperty(rNode, 'w:pict', 'pict');
  if (isRecord(pictNode)) {
    const rectNode = getProperty(pictNode, 'v:rect', 'rect');
    if (isRecord(rectNode)) {
      const hrAttr = getProperty(rectNode, '@_o:hr', '@_hr', '@o:hr', 'hr');
      if (hrAttr === 't' || hrAttr === true || hrAttr === 'true') {
        horizontalLine = true;
      }
    }
  }

  // Return run even if text is empty, as long as it was a valid run element
  const run: TextRun = {
    text,
    ...(bold ? { bold: true } : {}),
    ...(italic ? { italic: true } : {}),
    ...(underline ? { underline: true } : {}),
    ...(color ? { color } : {}),
    ...(highlight ? { highlight } : {}),
    ...(linkTarget ? { linkTarget } : {}),
    ...(pageBreak ? { pageBreak: true } : {}),
    ...(horizontalLine ? { horizontalLine: true } : {}),
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
 * @param defaultColor - Optional default text color inherited from paragraph properties.
 * @returns Array of parsed TextRun objects.
 */
function extractRunsFromContainer(
  container: Record<string, unknown>,
  imageMap?: Map<string, DocxImage>,
  paragraphImages?: DocxImage[],
  defaultColor?: string,
  hyperlinkMap?: Map<string, string>,
): TextRun[] {
  const runs: TextRun[] = [];

  // Direct runs (w:r)
  const directRuns = toArray(getProperty(container, 'w:r', 'r'));
  for (const rNode of directRuns) {
    const run = extractRun(rNode, imageMap, paragraphImages, defaultColor, hyperlinkMap);
    if (run !== null) {
      runs.push(run);
    }

    // Check for text boxes in drawings or pict elements inside this run
    if (isRecord(rNode)) {
      const drawingNodes = toArray(
        getProperty(rNode, 'w:drawing', 'drawing', 'w:pict', 'pict'),
      );
      for (const drawing of drawingNodes) {
        const txbxNodes = findTxbxContentNodes(drawing);
        for (const txbx of txbxNodes) {
          const pNodes = toArray(getProperty(txbx, 'w:p', 'p'));
          for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
            const p = pNodes[pIdx];
            if (isRecord(p)) {
              if (pIdx > 0 && runs.length > 0) {
                runs.push({ text: '\n' });
              }
              runs.push(...extractRunsFromContainer(p, imageMap, paragraphImages, undefined, hyperlinkMap));
            }
          }
        }
      }
    }
  }

  // Runs nested inside hyperlinks (w:hyperlink)
  const hyperlinks = toArray<Record<string, unknown>>(
    getProperty(container, 'w:hyperlink', 'hyperlink'),
  );
  for (const hyperlink of hyperlinks) {
    if (isRecord(hyperlink)) {
      runs.push(...extractRunsFromContainer(hyperlink, imageMap, paragraphImages, defaultColor, hyperlinkMap));
    }
  }

  // Also check container directly for drawing/pict nodes containing text boxes
  const directDrawings = toArray(
    getProperty(container, 'w:drawing', 'drawing', 'w:pict', 'pict'),
  );
  for (const drawing of directDrawings) {
    const txbxNodes = findTxbxContentNodes(drawing);
    for (const txbx of txbxNodes) {
      const pNodes = toArray(getProperty(txbx, 'w:p', 'p'));
      for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
        const p = pNodes[pIdx];
        if (isRecord(p)) {
          if (pIdx > 0 && runs.length > 0) {
            runs.push({ text: '\n' });
          }
          runs.push(...extractRunsFromContainer(p, imageMap, paragraphImages, undefined, hyperlinkMap));
        }
      }
    }
  }

  return runs;
}

/**
 * Extracts paragraph styles and their formatting properties (including alignment) from word/styles.xml.
 * Resolves basedOn inheritance chains so derived styles inherit parent alignment.
 *
 * @param stylesXml - Raw XML string from `word/styles.xml`.
 * @returns Map of style ID (and style name) to DocxStyleInfo.
 */
export function extractStyleMap(stylesXml: string): Map<string, DocxStyleInfo> {
  const map = new Map<string, DocxStyleInfo>();
  if (typeof stylesXml !== 'string' || stylesXml.trim().length === 0) {
    return map;
  }

  let parsed: Record<string, unknown>;
  try {
    parsed = parseXml(stylesXml);
  } catch {
    return map;
  }

  const stylesRoot = getProperty(parsed, 'w:styles', 'styles');
  const rootObj = isRecord(stylesRoot) ? stylesRoot : parsed;
  const styleNodes = toArray<Record<string, unknown>>(
    getProperty(rootObj, 'w:style', 'style'),
  );

  for (const styleNode of styleNodes) {
    if (!isRecord(styleNode)) {
      continue;
    }

    const type = getProperty(styleNode, '@_w:type', '@_type', '@w:type', 'type');
    const typeStr = type !== undefined ? String(type).toLowerCase().trim() : undefined;
    if (typeStr !== undefined && typeStr !== 'paragraph' && typeStr !== 'table') {
      continue;
    }

    const styleId = getProperty(
      styleNode,
      '@_w:styleId',
      '@_styleId',
      '@w:styleId',
      'styleId',
    );
    if (styleId === undefined || styleId === null) {
      continue;
    }
    const styleIdStr = String(styleId);

    let name: string | undefined;
    const nameNode = getProperty(styleNode, 'w:name', 'name');
    if (isRecord(nameNode)) {
      const val = getProperty(nameNode, '@_w:val', '@_val', '@w:val', 'val');
      if (val !== undefined && val !== null) {
        name = String(val);
      }
    }

    let basedOn: string | undefined;
    const basedOnNode = getProperty(styleNode, 'w:basedOn', 'basedOn');
    if (isRecord(basedOnNode)) {
      const val = getProperty(basedOnNode, '@_w:val', '@_val', '@w:val', 'val');
      if (val !== undefined && val !== null) {
        basedOn = String(val);
      }
    }

    let isDefault: boolean | undefined;
    const defaultVal = getProperty(
      styleNode,
      '@_w:default',
      '@_default',
      '@w:default',
      'default',
    );
    if (defaultVal === '1' || defaultVal === 1 || defaultVal === true || defaultVal === 'true') {
      isDefault = true;
    }

    let align: 'left' | 'center' | 'right' | 'justify' | undefined;
    let indent: ParagraphIndent | undefined;
    let spacing: DocxStyleInfo['spacing'] | undefined;
    const pPr = getProperty(styleNode, 'w:pPr', 'pPr');
    if (isRecord(pPr)) {
      const jcRaw = getProperty(pPr, 'w:jc', 'jc');
      let jcVal: string | undefined;
      if (isRecord(jcRaw)) {
        const val = getProperty(jcRaw, '@_w:val', '@_val', '@w:val', 'val');
        if (val !== undefined && val !== null) {
          jcVal = String(val).toLowerCase().trim();
        }
      } else if (typeof jcRaw === 'string') {
        jcVal = jcRaw.toLowerCase().trim();
      }

      if (jcVal === 'center') {
        align = 'center';
      } else if (jcVal === 'right') {
        align = 'right';
      } else if (jcVal === 'both' || jcVal === 'justify') {
        align = 'justify';
      } else if (jcVal === 'left') {
        align = 'left';
      }

      const indNode = getProperty(pPr, 'w:ind', 'ind');
      if (isRecord(indNode)) {
        const left = getNumericAttr(indNode, '@_w:left', '@_left', '@w:left', 'left', '@_w:start', '@w:start');
        const right = getNumericAttr(indNode, '@_w:right', '@_right', '@w:right', 'right', '@_w:end', '@w:end');
        const firstLine = getNumericAttr(indNode, '@_w:firstLine', '@_firstLine', '@w:firstLine', 'firstLine');
        const hanging = getNumericAttr(indNode, '@_w:hanging', '@_hanging', '@w:hanging', 'hanging');

        if (
          (left !== undefined && left > 0) ||
          (right !== undefined && right > 0) ||
          (firstLine !== undefined && firstLine > 0) ||
          (hanging !== undefined && hanging > 0)
        ) {
          indent = {};
          if (left !== undefined && left > 0) indent.left = left;
          if (right !== undefined && right > 0) indent.right = right;
          if (firstLine !== undefined && firstLine > 0) indent.firstLine = firstLine;
          if (hanging !== undefined && hanging > 0) indent.hanging = hanging;
        }
      }

      const spacingNode = getProperty(pPr, 'w:spacing', 'spacing');
      if (isRecord(spacingNode)) {
        const after = getNumericAttr(spacingNode, '@_w:after', '@_after', '@w:after', 'after');
        const before = getNumericAttr(spacingNode, '@_w:before', '@_before', '@w:before', 'before');
        const line = getNumericAttr(spacingNode, '@_w:line', '@_line', '@w:line', 'line');

        if (after !== undefined || before !== undefined || line !== undefined) {
          spacing = {};
          if (after !== undefined) spacing.after = after;
          if (before !== undefined) spacing.before = before;
          if (line !== undefined) spacing.line = line;
        }
      }
    }

    const info: DocxStyleInfo = {
      styleId: styleIdStr,
      ...(name !== undefined ? { name } : {}),
      ...(basedOn !== undefined ? { basedOn } : {}),
      ...(align !== undefined ? { align } : {}),
      ...(indent !== undefined ? { indent } : {}),
      ...(spacing !== undefined ? { spacing } : {}),
      ...(isDefault ? { isDefault: true } : {}),
    };

    map.set(styleIdStr, info);
    if (name && name !== styleIdStr && !map.has(name)) {
      map.set(name, info);
    }
  }

  // Resolve basedOn inheritance chains for align, indent, and spacing
  for (const info of map.values()) {
    if ((!info.align || !info.indent || !info.spacing) && info.basedOn) {
      let currentParentId: string | undefined = info.basedOn;
      const visited = new Set<string>([info.styleId]);
      while (currentParentId && !visited.has(currentParentId)) {
        visited.add(currentParentId);
        const parent = map.get(currentParentId);
        if (!parent) break;
        if (!info.align && parent.align) {
          info.align = parent.align;
        }
        if (!info.indent && parent.indent) {
          info.indent = { ...parent.indent };
        }
        if (!info.spacing && parent.spacing) {
          info.spacing = { ...parent.spacing };
        }
        if (info.align && info.indent && info.spacing) break;
        currentParentId = parent.basedOn;
      }
    }
  }

  return map;
}

/**
 * Extracts a Paragraph object from a parsed w:p node.
 *
 * @param pNode - Parsed w:p XML node.
 * @param options - Parser options.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @param styleMap - Optional mapping of style identifiers to DocxStyleInfo for style inheritance.
 * @returns Extracted Paragraph.
 */
export function extractParagraph(
  pNode: unknown,
  options?: DocxParserOptions,
  imageMap?: Map<string, DocxImage>,
  styleMap?: Map<string, DocxStyleInfo>,
  hyperlinkMap?: Map<string, string>,
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
        
        // If it's a heading, but the paragraph explicitly removes bold, downgrade it to a normal paragraph.
        // This handles cases where MS Word users abuse heading styles for spacing but unbold the text.
        if (/^heading/i.test(style)) {
          const prRPrRaw = getProperty(pPr, 'w:rPr', 'rPr');
          if (isRecord(prRPrRaw)) {
            const prBoldRaw = getProperty(prRPrRaw, 'w:b', 'b');
            if (prBoldRaw !== undefined) {
              const val = getProperty(prBoldRaw as any, '@_w:val', '@_val');
              const valStr = val !== undefined ? String(val).toLowerCase().trim() : '';
              if (valStr === 'false' || valStr === '0' || valStr === 'off') {
                style = undefined;
              }
            }
          }
        }
      }
    }
  }

  // Check for section break properties (w:sectPr / sectPr)
  let sectionBreak: SectionProperties | undefined;
  if (pPr) {
    const sectPrRaw = getProperty(pPr, 'w:sectPr', 'sectPr');
    if (isRecord(sectPrRaw)) {
      let columnCount = 1;
      const colsRaw = getProperty(sectPrRaw, 'w:cols', 'cols');
      if (isRecord(colsRaw)) {
        const numVal = getProperty(colsRaw, '@_w:num', '@_num', '@w:num', 'num', '@_w:val', '@_val', 'val');
        if (numVal !== undefined && numVal !== null) {
          const parsed = parseInt(String(numVal), 10);
          if (!isNaN(parsed)) {
            columnCount = parsed;
          }
        }
      }

      let sectPageBreak = false;
      const typeRaw = getProperty(sectPrRaw, 'w:type', 'type');
      let typeVal: string | undefined;
      if (isRecord(typeRaw)) {
        const val = getProperty(typeRaw, '@_w:val', '@_val', '@w:val', 'val');
        if (val !== undefined && val !== null) {
          typeVal = String(val).toLowerCase().trim();
        }
      } else if (typeof typeRaw === 'string') {
        typeVal = typeRaw.toLowerCase().trim();
      }

      if (!typeVal || typeVal === 'nextpage' || typeVal === 'oddpage' || typeVal === 'evenpage') {
        sectPageBreak = true;
      }

      if (columnCount > 1 || sectPageBreak || isRecord(colsRaw)) {
        sectionBreak = {
          columnCount,
          ...(sectPageBreak ? { pageBreak: true } : {}),
        };
      } else if (isRecord(sectPrRaw)) {
        // Even a 1-col section break (no cols element) is meaningful as a boundary marker
        sectionBreak = { columnCount: 1 };
      }
    }
  }

  // Check for paragraph default color in w:pPr > w:rPr > w:color
  let defaultColor: string | undefined;
  if (pPr) {
    const pRprRaw = getProperty(pPr, 'w:rPr', 'rPr');
    if (isRecord(pRprRaw)) {
      const colorNode = getProperty(pRprRaw, 'w:color', 'color');
      if (isRecord(colorNode)) {
        const colorVal = getProperty(colorNode, '@_w:val', '@_val', '@w:val', 'val');
        if (colorVal && typeof colorVal === 'string' && colorVal !== 'auto' && colorVal !== '000000') {
          defaultColor = colorVal;
        }
      }
    }
  }

  // Check for w:pageBreakBefore in pPr
  let pageBreakBefore: boolean | undefined;
  if (pPr) {
    const pbbNode = getProperty(pPr, 'w:pageBreakBefore', 'pageBreakBefore');
    if (pbbNode !== undefined && pbbNode !== null) {
      if (isRecord(pbbNode)) {
        const val = getProperty(pbbNode, '@_w:val', '@_val', '@w:val', 'val');
        if (val !== undefined && val !== null) {
          const valStr = String(val).toLowerCase().trim();
          if (valStr !== '0' && valStr !== 'false' && valStr !== 'off') {
            pageBreakBefore = true;
          }
        } else {
          pageBreakBefore = true;
        }
      } else if (typeof pbbNode === 'boolean') {
        if (pbbNode) pageBreakBefore = true;
      } else if (typeof pbbNode === 'number') {
        if (pbbNode !== 0) pageBreakBefore = true;
      } else if (typeof pbbNode === 'string') {
        const s = pbbNode.toLowerCase().trim();
        if (s !== '0' && s !== 'false' && s !== 'off') {
          pageBreakBefore = true;
        }
      } else {
        pageBreakBefore = true;
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

  // Check for paragraph alignment (w:jc / jc)
  let align: 'left' | 'center' | 'right' | 'justify' | undefined;
  if (pPr) {
    const jcRaw = getProperty(pPr, 'w:jc', 'jc');
    let jcVal: string | undefined;
    if (isRecord(jcRaw)) {
      const val = getProperty(jcRaw, '@_w:val', '@_val', '@w:val', 'val');
      if (val !== undefined && val !== null) {
        jcVal = String(val).toLowerCase().trim();
      }
    } else if (typeof jcRaw === 'string') {
      jcVal = jcRaw.toLowerCase().trim();
    }

    if (jcVal === 'center') {
      align = 'center';
    } else if (jcVal === 'right') {
      align = 'right';
    } else if (jcVal === 'both' || jcVal === 'justify') {
      align = 'justify';
    }
  }

  // Inherit alignment from style if not explicitly set on paragraph
  if (align === undefined && style !== undefined && styleMap) {
    const styleInfo = styleMap.get(style);
    if (styleInfo?.align) {
      align = styleInfo.align;
    }
  }

  // Check for paragraph indentation (w:ind / ind)
  let indent: ParagraphIndent | undefined;
  if (pPr) {
    const indNode = getProperty(pPr, 'w:ind', 'ind');
    if (isRecord(indNode)) {
      const left = getNumericAttr(indNode, '@_w:left', '@_left', '@w:left', 'left', '@_w:start', '@w:start');
      const right = getNumericAttr(indNode, '@_w:right', '@_right', '@w:right', 'right', '@_w:end', '@w:end');
      const firstLine = getNumericAttr(indNode, '@_w:firstLine', '@_firstLine', '@w:firstLine', 'firstLine');
      const hanging = getNumericAttr(indNode, '@_w:hanging', '@_hanging', '@w:hanging', 'hanging');

      // Only store if any value is non-zero
      if (
        (left !== undefined && left > 0) ||
        (right !== undefined && right > 0) ||
        (firstLine !== undefined && firstLine > 0) ||
        (hanging !== undefined && hanging > 0)
      ) {
        indent = {};
        if (left !== undefined && left > 0) indent.left = left;
        if (right !== undefined && right > 0) indent.right = right;
        if (firstLine !== undefined && firstLine > 0) indent.firstLine = firstLine;
        if (hanging !== undefined && hanging > 0) indent.hanging = hanging;
      }
    }
  }

  // Inherit indentation from style if not explicitly set on paragraph
  if (indent === undefined && style !== undefined && styleMap) {
    const styleInfo = styleMap.get(style);
    if (styleInfo?.indent) {
      indent = { ...styleInfo.indent };
    }
  }

  // Check for paragraph spacing (w:spacing / spacing)
  let spacing: Paragraph['spacing'] | undefined;
  if (pPr) {
    const spacingNode = getProperty(pPr, 'w:spacing', 'spacing');
    if (isRecord(spacingNode)) {
      const after = getNumericAttr(spacingNode, '@_w:after', '@_after', '@w:after', 'after');
      const before = getNumericAttr(spacingNode, '@_w:before', '@_before', '@w:before', 'before');
      const line = getNumericAttr(spacingNode, '@_w:line', '@_line', '@w:line', 'line');

      if (after !== undefined || before !== undefined || line !== undefined) {
        spacing = {};
        if (after !== undefined) spacing.after = after;
        if (before !== undefined) spacing.before = before;
        if (line !== undefined) spacing.line = line;
      }
    }
  }

  // Inherit spacing from style if not explicitly set on paragraph
  if (spacing === undefined && style !== undefined && styleMap) {
    const styleInfo = styleMap.get(style);
    if (styleInfo?.spacing) {
      spacing = { ...styleInfo.spacing };
    }
  }

  // Check for paragraph borders (w:pBdr / pBdr)
  let hasBorderTop: boolean | undefined;
  let hasBorderBottom: boolean | undefined;
  if (pPr) {
    const pBdrRaw = getProperty(pPr, 'w:pBdr', 'pBdr');
    if (isRecord(pBdrRaw)) {
      const bottomBdr = getProperty(pBdrRaw, 'w:bottom', 'bottom');
      if (isRecord(bottomBdr)) {
        const val = getProperty(bottomBdr, '@_w:val', '@_val', '@w:val', 'val');
        if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
          hasBorderBottom = true;
        }
      }
      const topBdr = getProperty(pBdrRaw, 'w:top', 'top');
      if (isRecord(topBdr)) {
        const val = getProperty(topBdr, '@_w:val', '@_val', '@w:val', 'val');
        if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
          hasBorderTop = true;
        }
      }
    }
  }

  const paragraphImages: DocxImage[] = [];
  const runs = extractRunsFromContainer(pNode, imageMap, paragraphImages, defaultColor, hyperlinkMap);
  if (defaultColor) {
    for (const run of runs) {
      if (!run.color && run.text && run.text.trim().length > 0) {
        run.color = defaultColor;
      }
    }
  }
  const text = runs.map((r) => r.text).join('');
  const hasPageBreakRun = runs.some((r) => r.pageBreak);
  const isEmpty =
    (runs.length === 0 || text.length === 0) &&
    paragraphImages.length === 0 &&
    !hasPageBreakRun;

  return {
    text,
    runs,
    ...(style !== undefined ? { style } : {}),
    ...(align !== undefined ? { align } : {}),
    ...(indent !== undefined ? { indent } : {}),
    ...(spacing !== undefined ? { spacing } : {}),
    ...(hasBorderTop ? { hasBorderTop: true } : {}),
    ...(hasBorderBottom ? { hasBorderBottom: true } : {}),
    ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
    ...(isEmpty ? { isEmpty: true } : {}),
    ...(listItem !== undefined ? { listItem } : {}),
    ...(paragraphImages.length > 0 ? { images: paragraphImages } : {}),
    ...(sectionBreak !== undefined ? { sectionBreak } : {}),
  };
}

/**
 * Parses raw WordprocessingML XML content into a structured DocxDocument.
 *
 * Navigates the OOXML hierarchy: `w:document` -> `w:body` -> `w:p` / `w:tbl` / `w:sdt`.
 * Handles formatting indicators (`w:b`, `w:i`, `w:u`), paragraph styles (`w:pStyle`),
 * multi-column sections, empty paragraphs, whitespace preservation, text boxes, and images.
 * Preserves document order via bodyItems.
 *
 * @param rawXml - Raw XML string from `word/document.xml`.
 * @param options - Optional parser configuration.
 * @param imageMap - Optional mapping of relationship IDs to DocxImage objects.
 * @returns The structured DocxDocument containing paragraphs, plain text, and body items.
 * @throws {DocxParseError} If the XML is malformed or cannot be parsed.
 */
export function extractText(
  rawXml: string,
  options?: DocxParserOptions,
  imageMap?: Map<string, DocxImage>,
  hyperlinkMap?: Map<string, string>,
): DocxDocument {
  const parsed = parseXml(rawXml);

  // Navigate to root document element
  const docRoot = getProperty(parsed, 'w:document', 'document');
  const rootObj = isRecord(docRoot) ? docRoot : parsed;

  // Navigate to body element
  const bodyRoot = getProperty(rootObj, 'w:body', 'body');
  const bodyObj = isRecord(bodyRoot) ? bodyRoot : rootObj;

  // Check for w:sectPr directly in bodyObj (final section properties)
  const sections: SectionProperties[] = [];
  if (isRecord(bodyObj)) {
    const bodySectPr = getProperty(bodyObj, 'w:sectPr', 'sectPr');
    if (isRecord(bodySectPr)) {
      let columnCount = 1;
      const colsNode = getProperty(bodySectPr, 'w:cols', 'cols');
      if (isRecord(colsNode)) {
        const numVal = getProperty(colsNode, '@_w:num', '@_num', '@w:num', 'num', '@_w:val', '@_val', 'val');
        if (numVal !== undefined && numVal !== null) {
          const parsed = parseInt(String(numVal), 10);
          if (!isNaN(parsed)) {
            columnCount = parsed;
          }
        }
      }
      if (columnCount > 1) {
        sections.push({ columnCount });
      }
    }
  }

  const pNodes = toArray(getProperty(bodyObj, 'w:p', 'p'));
  const tblNodes = toArray<Record<string, unknown>>(getProperty(bodyObj, 'w:tbl', 'tbl'));
  const sdtNodes = toArray<Record<string, unknown>>(getProperty(bodyObj, 'w:sdt', 'sdt'));

  const styleMap = options?.stylesXml ? extractStyleMap(options.stylesXml) : undefined;

  let orderedParagraphs: Paragraph[] = [];
  let orderedTables: DocxTable[] = [];
  let bodyItems: BodyItem[] = [];

  let preserveOrderSuccess = false;

  try {
    const orderedRoots = parseXmlPreserveOrder(rawXml);
    const findChildInOrdered = (
      items: Array<Record<string, unknown>>,
      ...names: string[]
    ): Array<Record<string, unknown>> | undefined => {
      for (const item of items) {
        for (const name of names) {
          if (name in item && Array.isArray(item[name])) {
            return item[name] as Array<Record<string, unknown>>;
          }
        }
      }
      return undefined;
    };

    const docChildren = findChildInOrdered(orderedRoots, 'w:document', 'document') ?? orderedRoots;
    const bodyChildren = findChildInOrdered(docChildren, 'w:body', 'body');

    if (bodyChildren && bodyChildren.length > 0) {
      let pIndex = 0;
      let tblIndex = 0;
      let sdtIndex = 0;

      for (const child of bodyChildren) {
        const tagName = Object.keys(child).find((k) => k !== ':@');
        if (!tagName) continue;

        if (tagName === 'w:p' || tagName === 'p') {
          if (pIndex < pNodes.length) {
            const para = extractParagraph(pNodes[pIndex++], options, imageMap, styleMap, hyperlinkMap);
            orderedParagraphs.push(para);
            bodyItems.push({ type: 'paragraph', paragraph: para });
          }
        } else if (tagName === 'w:tbl' || tagName === 'tbl') {
          if (tblIndex < tblNodes.length) {
            const tableWrapper = { 'w:tbl': [tblNodes[tblIndex++]] };
            const extracted = extractTables(tableWrapper, options, imageMap, styleMap);
            if (extracted.length > 0) {
              orderedTables.push(extracted[0]);
              bodyItems.push({ type: 'table', table: extracted[0] });
            }
          }
        } else if (tagName === 'w:sdt' || tagName === 'sdt') {
          if (sdtIndex < sdtNodes.length) {
            const sdtNode = sdtNodes[sdtIndex++];
            const sdtContent = getProperty(sdtNode, 'w:sdtContent', 'sdtContent');
            if (isRecord(sdtContent)) {
              const sdtPNodes = toArray(getProperty(sdtContent, 'w:p', 'p'));
              for (const sdtP of sdtPNodes) {
                const para = extractParagraph(sdtP, options, imageMap, styleMap, hyperlinkMap);
                orderedParagraphs.push(para);
                bodyItems.push({ type: 'paragraph', paragraph: para });
              }
              const sdtTblNodes = toArray(getProperty(sdtContent, 'w:tbl', 'tbl'));
              for (const sdtTbl of sdtTblNodes) {
                const tableWrapper = { 'w:tbl': [sdtTbl] };
                const extracted = extractTables(tableWrapper, options, imageMap, styleMap);
                if (extracted.length > 0) {
                  orderedTables.push(extracted[0]);
                  bodyItems.push({ type: 'table', table: extracted[0] });
                }
              }
            }
          }
        }
      }

      // Append any remaining paragraphs or tables if not visited in bodyChildren
      while (pIndex < pNodes.length) {
        const para = extractParagraph(pNodes[pIndex++], options, imageMap, styleMap, hyperlinkMap);
        orderedParagraphs.push(para);
        bodyItems.push({ type: 'paragraph', paragraph: para });
      }
      while (tblIndex < tblNodes.length) {
        const tableWrapper = { 'w:tbl': [tblNodes[tblIndex++]] };
        const extracted = extractTables(tableWrapper, options, imageMap, styleMap);
        if (extracted.length > 0) {
          orderedTables.push(extracted[0]);
          bodyItems.push({ type: 'table', table: extracted[0] });
        }
      }

      preserveOrderSuccess = true;
    }
  } catch {
    // Graceful fallback
  }

  if (!preserveOrderSuccess) {
    orderedParagraphs = pNodes.map((pNode) =>
      extractParagraph(pNode, options, imageMap, styleMap, hyperlinkMap),
    );
    orderedTables = isRecord(bodyObj) ? extractTables(bodyObj, options, imageMap, styleMap) : [];
    bodyItems = [
      ...orderedParagraphs.map((p) => ({ type: 'paragraph' as const, paragraph: p })),
      ...orderedTables.map((t) => ({ type: 'table' as const, table: t })),
    ];
  }



  const text = orderedParagraphs.map((p) => p.text).join('\n');

  return {
    paragraphs: orderedParagraphs,
    text,
    tables: orderedTables,
    ...(sections.length > 0 ? { sections } : {}),
    bodyItems,
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
 * @param abstractNumFmtMap - Optional mapping from "${abstractNumId}:${level}" to numFmt string.
 * @returns A new DocxDocument with resolved list items.
 */
export function resolveListItems(
  doc: DocxDocument,
  numIdMap: NumIdMap,
  abstractNumMap: AbstractNumMap,
  abstractNumFmtMap?: AbstractNumFmtMap,
): DocxDocument {
  const resolvePara = (para: Paragraph): Paragraph => {
    if (!para.listItem) {
      return { ...para };
    }

    const { numId, level } = para.listItem;
    let abstractNumId = numIdMap.get(numId);
    if (abstractNumId === undefined) {
      if (abstractNumMap.has(`${numId}:${level}`) || abstractNumMap.has(`${numId}:0`)) {
        abstractNumId = numId;
      } else {
        abstractNumId = 0;
      }
    }

    let listType = abstractNumMap.get(`${abstractNumId}:${level}`);
    let numFmt = abstractNumFmtMap?.get(`${abstractNumId}:${level}`);

    if (listType === undefined) {
      for (let l = level - 1; l >= 0; l--) {
        const fallbackType = abstractNumMap.get(`${abstractNumId}:${l}`);
        if (fallbackType !== undefined) {
          listType = fallbackType;
          numFmt = abstractNumFmtMap?.get(`${abstractNumId}:${l}`);
          break;
        }
      }
    }

    if (listType === undefined) {
      for (const [key, type] of abstractNumMap.entries()) {
        if (key.startsWith(`${abstractNumId}:`)) {
          listType = type;
          numFmt = abstractNumFmtMap?.get(key);
          break;
        }
      }
    }

    if (listType === undefined) {
      listType = 'bullet';
    }

    const listItem: ListItemInfo = {
      abstractNumId,
      numId,
      level,
      listType,
      ...(numFmt !== undefined ? { numFmt } : {}),
    };

    return {
      ...para,
      listItem,
    };
  };

  const paragraphs: Paragraph[] = doc.paragraphs.map(resolvePara);

  const bodyItems = doc.bodyItems?.map((item) => {
    if (item.type === 'paragraph') {
      return {
        type: 'paragraph' as const,
        paragraph: resolvePara(item.paragraph),
      };
    }
    return item;
  });

  return {
    ...doc,
    paragraphs,
    ...(bodyItems ? { bodyItems } : {}),
  };
}
