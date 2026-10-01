/**
 * @file tableExtractor.ts
 * Module for extracting structured tables (w:tbl) from OOXML WordprocessingML content.
 */

import { extractParagraph } from './textExtractor.js';
import type {
  DocxImage,
  DocxParserOptions,
  DocxStyleInfo,
  DocxTable,
  Paragraph,
  TableCell,
  TableRow,
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
 * Checks whether a TableCell contains no meaningful text or images.
 *
 * @param cell - Target table cell.
 * @returns True if the cell is completely empty.
 */
function isCellEmpty(cell: TableCell): boolean {
  if (!cell.paragraphs || cell.paragraphs.length === 0) {
    return true;
  }
  return cell.paragraphs.every((p) => {
    const hasText = p.text && p.text.trim().length > 0;
    const hasImages = Array.isArray(p.images) && p.images.length > 0;
    return !hasText && !hasImages;
  });
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
 * Extracts all tables (`w:tbl`) from a parsed WordprocessingML body node.
 *
 * Navigates the OOXML hierarchy: `w:tbl` -> `w:tr` -> `w:tc` -> `w:p`.
 * Extracts column count, row structure, cell spans, vertical merge indicators,
 * and detects header rows based on paragraph styles.
 *
 * @param bodyNode - The parsed `w:body` node from the document XML tree.
 * @returns An array of structured DocxTable objects.
 */
export function extractTables(
  bodyNode: Record<string, unknown>,
  options?: DocxParserOptions,
  imageMap?: Map<string, DocxImage>,
  styleMap?: Map<string, DocxStyleInfo>,
): DocxTable[] {
  if (!isRecord(bodyNode)) {
    return [];
  }

  const tblNodes = toArray<Record<string, unknown>>(getProperty(bodyNode, 'w:tbl', 'tbl'));
  const tables: DocxTable[] = [];

  for (const tblNode of tblNodes) {
    if (!isRecord(tblNode)) {
      continue;
    }

    // Extract table style from w:tblPr > w:tblStyle @w:val
    let style: string | undefined;
    const tblPr = getProperty(tblNode, 'w:tblPr', 'tblPr');
    if (isRecord(tblPr)) {
      const tblStyle = getProperty(tblPr, 'w:tblStyle', 'tblStyle');
      if (isRecord(tblStyle)) {
        const val = getProperty(tblStyle, '@_w:val', '@_val', '@w:val', 'val');
        if (val !== undefined && val !== null) {
          style = String(val);
        }
      } else if (typeof tblStyle === 'string') {
        style = tblStyle;
      }
    }

    // Extract table borders from w:tblPr > w:tblBorders
    let borders: DocxTable['borders'] | undefined;
    if (isRecord(tblPr)) {
      const tblBordersNode = getProperty(tblPr, 'w:tblBorders', 'tblBorders');
      if (isRecord(tblBordersNode)) {
        const checkBorder = (nodeKey: string, aliasKey: string): boolean => {
          const borderNode = getProperty(tblBordersNode, nodeKey, aliasKey);
          if (isRecord(borderNode)) {
            const val = getProperty(borderNode, '@_w:val', '@_val', '@w:val', 'val');
            if (val !== undefined && val !== null) {
              const valStr = String(val).toLowerCase().trim();
              return valStr !== 'none' && valStr !== 'nil';
            }
          }
          return false;
        };

        borders = {
          top: checkBorder('w:top', 'top'),
          bottom: checkBorder('w:bottom', 'bottom'),
          left: checkBorder('w:left', 'left'),
          right: checkBorder('w:right', 'right'),
          insideH: checkBorder('w:insideH', 'insideH'),
          insideV: checkBorder('w:insideV', 'insideV'),
        };
      }
    }

    // Rows in table
    const trNodes = toArray<Record<string, unknown>>(getProperty(tblNode, 'w:tr', 'tr'));
    let hasAnyCellBorders = false;
    let hasCellExplicitlyRemovedVertical = false;
    let hasAnyCellVerticalBorder = false;

    // Determine columnCount: from w:tblGrid > w:gridCol OR first row gridSpan sum
    let columnCount = 0;
    const tblGrid = getProperty(tblNode, 'w:tblGrid', 'tblGrid');
    if (isRecord(tblGrid)) {
      const gridCols = toArray(getProperty(tblGrid, 'w:gridCol', 'gridCol'));
      if (gridCols.length > 0) {
        columnCount = gridCols.length;
      }
    }

    if (columnCount === 0 && trNodes.length > 0) {
      for (const r of trNodes) {
        if (isRecord(r)) {
          const tcList = toArray<Record<string, unknown>>(getProperty(r, 'w:tc', 'tc'));
          let rowCols = 0;
          for (const tc of tcList) {
            let span = 1;
            if (isRecord(tc)) {
              const tcPr = getProperty(tc, 'w:tcPr', 'tcPr');
              if (isRecord(tcPr)) {
                const gs = getProperty(tcPr, 'w:gridSpan', 'gridSpan');
                if (isRecord(gs)) {
                  const val = getProperty(gs, '@_w:val', '@_val', '@w:val', 'val');
                  const parsed = parseInt(String(val), 10);
                  if (!isNaN(parsed)) span = parsed;
                } else if (typeof gs === 'number') {
                  span = Math.trunc(gs);
                } else if (typeof gs === 'string') {
                  const parsed = parseInt(gs, 10);
                  if (!isNaN(parsed)) span = parsed;
                }
              }
            }
            rowCols += span;
          }
          if (rowCols > 0) {
            columnCount = rowCols;
            break;
          }
        }
      }
    }

    // Extract column widths from w:tblGrid > w:gridCol
    let columnWidths: number[] | undefined;
    if (isRecord(tblGrid)) {
      const gridCols = toArray(getProperty(tblGrid, 'w:gridCol', 'gridCol'));
      const widths: number[] = [];
      for (const gc of gridCols) {
        if (isRecord(gc)) {
          const w = getProperty(gc, '@_w:w', '@_w', '@w:w', 'w');
          if (w !== undefined && w !== null) {
            const parsed = parseInt(String(w), 10);
            widths.push(!isNaN(parsed) ? parsed : 0);
          } else {
            widths.push(0);
          }
        }
      }
      if (widths.length > 0) {
        columnWidths = widths;
      }
    }

    const rows: TableRow[] = [];

    for (const trNode of trNodes) {
      if (!isRecord(trNode)) {
        continue;
      }

      const tcNodes = toArray<Record<string, unknown>>(getProperty(trNode, 'w:tc', 'tc'));
      const cells: TableCell[] = [];
      let consumed = 0;

      for (const tcNode of tcNodes) {
        if (!isRecord(tcNode)) {
          continue;
        }

        const tcPr = getProperty(tcNode, 'w:tcPr', 'tcPr');
        let columnSpan: number | undefined;
        let isVerticalMerge: boolean | undefined;
        let vMerge: 'restart' | 'continue' | undefined;
        let cellBorders: { top?: boolean; bottom?: boolean; left?: boolean; right?: boolean } | undefined;

        if (isRecord(tcPr)) {
          // Extract columnSpan from w:tcPr > w:gridSpan @w:val (default 1 if missing)
          const gridSpanNode = getProperty(tcPr, 'w:gridSpan', 'gridSpan');
          if (gridSpanNode !== undefined) {
            if (isRecord(gridSpanNode)) {
              const val = getProperty(gridSpanNode, '@_w:val', '@_val', '@w:val', 'val');
              if (val !== undefined && val !== null) {
                const parsed = parseInt(String(val), 10);
                columnSpan = !isNaN(parsed) ? parsed : 1;
              } else {
                columnSpan = 1;
              }
            } else if (typeof gridSpanNode === 'number') {
              columnSpan = Math.trunc(gridSpanNode);
            } else if (typeof gridSpanNode === 'string') {
              const parsed = parseInt(gridSpanNode, 10);
              columnSpan = !isNaN(parsed) ? parsed : 1;
            } else {
              columnSpan = 1;
            }
          }

          // Extract isVerticalMerge: if w:tcPr > w:vMerge element exists
          const vMergeNode = getProperty(tcPr, 'w:vMerge', 'vMerge');
          if (vMergeNode !== undefined) {
            isVerticalMerge = true;
            let val: string | undefined;
            if (isRecord(vMergeNode)) {
              const rawVal = getProperty(vMergeNode, '@_w:val', '@_val', '@w:val', 'val');
              if (rawVal !== undefined && rawVal !== null) {
                val = String(rawVal).toLowerCase().trim();
              }
            } else if (typeof vMergeNode === 'string') {
              val = vMergeNode.toLowerCase().trim();
            }

            if (val === 'restart') {
              vMerge = 'restart';
            } else {
              vMerge = 'continue';
            }
          }

          // Check cell borders
          const tcBorders = getProperty(tcPr, 'w:tcBorders', 'tcBorders');
          if (isRecord(tcBorders)) {
            hasAnyCellBorders = true;
            cellBorders = {};
            for (const side of ['top', 'bottom', 'left', 'right'] as const) {
              const bNode = getProperty(tcBorders, `w:${side}`, side);
              if (isRecord(bNode)) {
                const val = getProperty(bNode, '@_w:val', '@_val', '@w:val', 'val');
                if (val !== undefined && val !== null) {
                  const s = String(val).toLowerCase().trim();
                  if (s === 'none' || s === 'nil') {
                    cellBorders[side] = false;
                    if (side === 'left' || side === 'right') hasCellExplicitlyRemovedVertical = true;
                  } else {
                    cellBorders[side] = true;
                    if (side === 'left' || side === 'right') hasAnyCellVerticalBorder = true;
                  }
                }
              }
            }
            if (Object.keys(cellBorders).length === 0) {
              cellBorders = undefined;
            }
          }
        }

        // Extract cell paragraphs
        const pNodes = toArray(getProperty(tcNode, 'w:p', 'p'));
        const paragraphs: Paragraph[] = pNodes.map((pNode) =>
          extractParagraph(pNode, options, imageMap, styleMap),
        );

        // Inherit table style spacing for cell paragraphs if not explicitly defined
        const tableStyleInfo = style ? styleMap?.get(style) : undefined;
        if (tableStyleInfo?.spacing) {
          for (const p of paragraphs) {
            if (!p.spacing) {
              p.spacing = { ...tableStyleInfo.spacing };
            }
          }
        }

        // Extract vertical alignment from w:tcPr > w:vAlign @w:val
        let align: 'top' | 'horizon' | 'bottom' | undefined;
        if (isRecord(tcPr)) {
          const vAlignNode = getProperty(tcPr, 'w:vAlign', 'vAlign');
          if (vAlignNode !== undefined) {
            let vAlignVal: string | undefined;
            if (isRecord(vAlignNode)) {
              const val = getProperty(vAlignNode, '@_w:val', '@_val', '@w:val', 'val');
              if (val !== undefined && val !== null) {
                vAlignVal = String(val).toLowerCase().trim();
              }
            } else if (typeof vAlignNode === 'string') {
              vAlignVal = vAlignNode.toLowerCase().trim();
            }

            if (vAlignVal === 'center') {
              align = 'horizon';
            } else if (vAlignVal === 'top') {
              align = 'top';
            } else if (vAlignVal === 'bottom') {
              align = 'bottom';
            }
          }
        }

        const cell: TableCell = {
          paragraphs,
          ...(columnSpan !== undefined ? { columnSpan } : {}),
          ...(isVerticalMerge ? { isVerticalMerge: true } : {}),
          ...(vMerge !== undefined ? { vMerge } : {}),
          ...(align !== undefined ? { align } : {}),
          ...(cellBorders !== undefined ? { borders: cellBorders } : {}),
        };

        // If logical cells already satisfy columnCount, skip extra phantom cells
        if (columnCount > 0 && consumed >= columnCount && isCellEmpty(cell)) {
          break;
        }

        cells.push(cell);

        consumed += columnSpan ?? 1;
      }

      // Drop trailing phantom cells (isVerticalMerge=true and isEmpty) if exceeding columnCount
      if (columnCount > 0) {
        while (cells.length > 0) {
          const totalSlots = cells.reduce((sum, c) => sum + (c.columnSpan ?? 1), 0);
          if (totalSlots <= columnCount) {
            break;
          }
          const lastCell = cells[cells.length - 1];
          if (lastCell.isVerticalMerge && isCellEmpty(lastCell)) {
            cells.pop();
          } else {
            break;
          }
        }
      }

      // Detect isHeader for the row: if any cell's first paragraph has a style containing 'heading' or 'header'
      let isHeader = false;
      for (const cell of cells) {
        const firstPara = cell.paragraphs[0];
        if (firstPara?.style && /heading|header/i.test(firstPara.style)) {
          isHeader = true;
          break;
        }
      }

      const row: TableRow = {
        cells,
        ...(isHeader ? { isHeader: true } : {}),
      };
      rows.push(row);
    }

    // Calculate rowSpan for vMerge='restart' cells
    for (let r = 0; r < rows.length; r++) {
      let colIdx = 0;
      for (const cell of rows[r].cells) {
        const span = cell.columnSpan ?? 1;
        if (cell.vMerge === 'restart') {
          let rowSpan = 1;
          for (let nextR = r + 1; nextR < rows.length; nextR++) {
            let nextCol = 0;
            let matchingCell: TableCell | undefined;
            for (const c of rows[nextR].cells) {
              if (nextCol === colIdx) {
                matchingCell = c;
                break;
              }
              nextCol += c.columnSpan ?? 1;
            }
            if (matchingCell && matchingCell.vMerge === 'continue') {
              rowSpan++;
            } else {
              break;
            }
          }
          if (rowSpan > 1) {
            cell.rowSpan = rowSpan;
          }
        }
        colIdx += span;
      }
    }

    let hasBorderTop = false;
    let hasBorderBottom = false;

    if (rows.length > 0) {
      const firstRow = rows[0];
      if (firstRow.cells.some((cell) => cell.paragraphs.some((p) => p.hasBorderTop))) {
        hasBorderTop = true;
      }
    }

    if (hasBorderTop && trNodes.length > 0) {
      const lastTrNode = trNodes[trNodes.length - 1];
      if (isRecord(lastTrNode)) {
        const lastTcNodes = toArray<Record<string, unknown>>(
          getProperty(lastTrNode, 'w:tc', 'tc'),
        );
        for (const tc of lastTcNodes) {
          if (isRecord(tc)) {
            const tcPr = getProperty(tc, 'w:tcPr', 'tcPr');
            if (isRecord(tcPr)) {
              const tcBorders = getProperty(tcPr, 'w:tcBorders', 'tcBorders');
              if (isRecord(tcBorders)) {
                const bottomBdr = getProperty(tcBorders, 'w:bottom', 'bottom');
                if (isRecord(bottomBdr)) {
                  const val = getProperty(bottomBdr, '@_w:val', '@_val', '@w:val', 'val');
                  if (typeof val === 'string' && val.toLowerCase() !== 'none' && val.toLowerCase() !== 'nil') {
                    hasBorderBottom = true;
                    break;
                  }
                }
              }
            }
          }
        }
      }
      if (!hasBorderBottom && rows.length > 0) {
        const lastRow = rows[rows.length - 1];
        if (lastRow.cells.some((cell) => cell.paragraphs.some((p) => p.hasBorderBottom))) {
          hasBorderBottom = true;
        }
      }
    }

    if (!borders && hasAnyCellBorders) {
      borders = {
        top: hasBorderTop,
        bottom: hasBorderBottom,
        left: hasAnyCellVerticalBorder,
        right: hasAnyCellVerticalBorder,
        insideH: false,
        insideV: hasAnyCellVerticalBorder,
      };
    }

    const table: DocxTable = {
      rows,
      columnCount,
      ...(columnWidths && columnWidths.length === columnCount ? { columnWidths } : {}),
      ...(style !== undefined ? { style } : {}),
      ...(borders !== undefined ? { borders } : {}),
      ...(hasBorderTop ? { hasBorderTop: true } : {}),
      ...(hasBorderBottom ? { hasBorderBottom: true } : {}),
    };
    tables.push(table);
  }

  return tables;
}
