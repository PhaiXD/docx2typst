/**
 * @file tableExtractor.ts
 * Module for extracting structured tables (w:tbl) from OOXML WordprocessingML content.
 */

import { extractParagraph } from './textExtractor.js';
import type { DocxTable, Paragraph, TableCell, TableRow } from '../types.js';

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
export function extractTables(bodyNode: Record<string, unknown>): DocxTable[] {
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

    // Rows in table
    const trNodes = toArray<Record<string, unknown>>(getProperty(tblNode, 'w:tr', 'tr'));

    // Determine columnCount: count w:tc in first row OR w:gridCol in w:tblGrid
    let columnCount = 0;
    const firstRow = trNodes.length > 0 && isRecord(trNodes[0]) ? trNodes[0] : undefined;
    if (firstRow) {
      const firstRowCells = toArray(getProperty(firstRow, 'w:tc', 'tc'));
      if (firstRowCells.length > 0) {
        columnCount = firstRowCells.length;
      }
    }
    if (columnCount === 0) {
      const tblGrid = getProperty(tblNode, 'w:tblGrid', 'tblGrid');
      if (isRecord(tblGrid)) {
        const gridCols = toArray(getProperty(tblGrid, 'w:gridCol', 'gridCol'));
        if (gridCols.length > 0) {
          columnCount = gridCols.length;
        }
      }
    }
    if (columnCount === 0 && trNodes.length > 0) {
      for (const r of trNodes) {
        if (isRecord(r)) {
          const count = toArray(getProperty(r, 'w:tc', 'tc')).length;
          if (count > columnCount) {
            columnCount = count;
          }
        }
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
          }
        }

        // Extract cell paragraphs
        const pNodes = toArray(getProperty(tcNode, 'w:p', 'p'));
        const paragraphs: Paragraph[] = pNodes.map((pNode) => extractParagraph(pNode));

        // Extract text alignment from w:tcPr > w:jc @w:val OR from first paragraph's w:pPr > w:jc @w:val
        let align: 'left' | 'center' | 'right' | undefined;
        let jcNode: unknown;
        if (isRecord(tcPr)) {
          jcNode = getProperty(tcPr, 'w:jc', 'jc');
        }
        if (jcNode === undefined && pNodes.length > 0 && isRecord(pNodes[0])) {
          const firstPPr = getProperty(pNodes[0], 'w:pPr', 'pPr');
          if (isRecord(firstPPr)) {
            jcNode = getProperty(firstPPr, 'w:jc', 'jc');
          }
        }

        if (jcNode !== undefined) {
          let jcVal: string | undefined;
          if (isRecord(jcNode)) {
            const val = getProperty(jcNode, '@_w:val', '@_val', '@w:val', 'val');
            if (val !== undefined && val !== null) {
              jcVal = String(val).toLowerCase().trim();
            }
          } else if (typeof jcNode === 'string') {
            jcVal = jcNode.toLowerCase().trim();
          }

          if (jcVal !== undefined) {
            if (jcVal === 'center') {
              align = 'center';
            } else if (jcVal === 'right') {
              align = 'right';
            } else if (jcVal === 'left' || jcVal === 'both' || jcVal === 'justify') {
              align = 'left';
            } else {
              align = 'left';
            }
          }
        }

        const cell: TableCell = {
          paragraphs,
          ...(columnSpan !== undefined ? { columnSpan } : {}),
          ...(isVerticalMerge ? { isVerticalMerge: true } : {}),
          ...(align !== undefined ? { align } : {}),
        };

        // If logical cells already satisfy columnCount, skip extra phantom cells
        if (columnCount > 0 && consumed >= columnCount && isCellEmpty(cell)) {
          break;
        }

        cells.push(cell);

        if (!isVerticalMerge) {
          consumed += columnSpan ?? 1;
        }
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

    // Detect and remove narrow spacer columns (BUG 3)
    const spacerCols: number[] = [];
    for (let c = 0; c < columnCount; c++) {
      let isSpacer = rows.length > 0;
      for (const row of rows) {
        let colIdx = 0;
        let cellAtCol: TableCell | undefined;
        let spanAtCol = 1;
        for (const cell of row.cells) {
          const span = cell.columnSpan ?? 1;
          if (colIdx <= c && c < colIdx + span) {
            cellAtCol = cell;
            spanAtCol = span;
            break;
          }
          colIdx += span;
        }
        if (!cellAtCol || spanAtCol > 1 || !isCellEmpty(cellAtCol)) {
          isSpacer = false;
          break;
        }
      }
      if (isSpacer) {
        spacerCols.push(c);
      }
    }

    if (spacerCols.length > 0) {
      for (let s = spacerCols.length - 1; s >= 0; s--) {
        const spacerCol = spacerCols[s];
        const newColCount = columnCount - 1;
        for (const row of rows) {
          let colIdx = 0;
          let cellIndex = -1;
          for (let i = 0; i < row.cells.length; i++) {
            const cell = row.cells[i];
            const span = cell.columnSpan ?? 1;
            if (colIdx <= spacerCol && spacerCol < colIdx + span) {
              cellIndex = i;
              break;
            }
            colIdx += span;
          }
          if (cellIndex !== -1) {
            const spacerCell = row.cells[cellIndex];
            row.cells.splice(cellIndex, 1);
            if (cellIndex > 0) {
              const prevCell = row.cells[cellIndex - 1];
              const spacerSpan = spacerCell.columnSpan ?? 1;
              const currentTotalSlots = row.cells.reduce((sum, cl) => sum + (cl.columnSpan ?? 1), 0);
              if (currentTotalSlots + spacerSpan <= newColCount) {
                prevCell.columnSpan = (prevCell.columnSpan ?? 1) + spacerSpan;
              }
            }
          }
        }
        columnCount = newColCount;
      }
    }

    const table: DocxTable = {
      rows,
      columnCount,
      ...(style !== undefined ? { style } : {}),
    };
    tables.push(table);
  }

  return tables;
}
