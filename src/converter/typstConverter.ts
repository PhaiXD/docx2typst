/**
 * @file typstConverter.ts
 * Converts structured DocxDocument content into Typst markup format.
 */

import { basename } from 'node:path';
import type {
  BodyItem,
  DocxDocument,
  DocxImage,
  DocxTable,
  Paragraph,
  TableCell,
  TextRun,
  TypstConverterOptions,
  TypstDocument,
} from '../types.js';

/**
 * Escapes Typst special characters in plain text portions.
 * Special characters escaped: @, #, <, >, ` (backtick), \ (backslash).
 *
 * @param text - Plain text string to escape.
 * @returns Escaped text string suitable for Typst markup.
 */
function escapeTypstText(text: string): string {
  return text.replace(/[\\#@<>`]/g, (char) => `\\${char}`);
}

/**
 * Maps an OOXML style identifier to a Typst heading level (1 through 6).
 * Recognizes patterns such as 'Heading1', 'heading1', 'Heading 1', '1', up to level 6.
 *
 * @param style - The paragraph style identifier to evaluate.
 * @returns Heading level number (1-6) or null if the style is not a recognized heading.
 */
export function getHeadingLevel(style: string | undefined): number | null {
  if (!style) {
    return null;
  }

  const trimmed = style.trim();

  // Match 'Heading1', 'heading1', 'Heading 1', 'heading 1', etc.
  const headingMatch = /^heading\s*([1-6])$/i.exec(trimmed);
  if (headingMatch) {
    return parseInt(headingMatch[1], 10);
  }

  // Match standalone digits '1' through '6'
  const digitMatch = /^([1-6])$/.exec(trimmed);
  if (digitMatch) {
    return parseInt(digitMatch[1], 10);
  }

  return null;
}

/**
 * Converts a single TextRun into Typst markup, applying bold, italic,
 * and underline formatting in the proper nesting order.
 *
 * Nesting order: underline (outermost) -> bold -> italic (innermost).
 *
 * @param run - The text run to convert.
 * @param escape - Whether to escape special Typst characters (default: true).
 * @returns The converted Typst markup string for this run.
 */
function convertRunToTypst(run: TextRun, escape: boolean = true): string {
  if (run.pageBreak) {
    const rawText = run.text || '';
    if (rawText.trim().length === 0) {
      return '#pagebreak()';
    }
    const text = escape ? escapeTypstText(rawText) : rawText;
    const match = text.match(/^(\s*)(.*?)(\s*)$/s);
    const leadingWs = match ? match[1] : '';
    const coreText = match ? match[2] : text;
    const trailingWs = match ? match[3] : '';

    let formatted = coreText;
    if (run.italic) {
      formatted = `_${formatted}_`;
    }
    if (run.bold) {
      formatted = `*${formatted}*`;
    }
    if (run.underline) {
      formatted = `#underline[${formatted}]`;
    }
    if (run.color) {
      formatted = `#text(fill: rgb("${run.color}"))[${formatted}]`;
    }
    return `#pagebreak()\n${leadingWs}${formatted}${trailingWs}`;
  }

  if (!run.text || run.text.length === 0) {
    return '';
  }

  const text = escape ? escapeTypstText(run.text) : run.text;

  // If no formatting is active, return the text directly
  if (!run.bold && !run.italic && !run.underline && !run.color) {
    return text;
  }

  // If text is purely whitespace, formatting markers cannot attach to non-whitespace
  if (text.trim().length === 0) {
    return text;
  }

  // Extract leading and trailing whitespace to keep formatting markers touching the word
  const match = text.match(/^(\s*)(.*?)(\s*)$/s);
  const leadingWs = match ? match[1] : '';
  const coreText = match ? match[2] : text;
  const trailingWs = match ? match[3] : '';

  let formatted = coreText;

  // Italic is innermost
  if (run.italic) {
    formatted = `_${formatted}_`;
  }

  // Bold wraps around italic
  if (run.bold) {
    formatted = `*${formatted}*`;
  }

  // Underline is outermost
  if (run.underline) {
    formatted = `#underline[${formatted}]`;
  }

  // Color wraps around underline
  if (run.color) {
    formatted = `#text(fill: rgb("${run.color}"))[${formatted}]`;
  }

  return `${leadingWs}${formatted}${trailingWs}`;
}

/**
 * Converts a DocxImage into Typst `#image(...)` markup.
 *
 * Emits `#image("path/to/image.png")` or `#image("path/to/image.png", width: XX%)`.
 * Width is calculated as a percentage of page width assuming 6 inches = 5,486,400 EMUs,
 * rounded to the nearest integer percent and capped at 100%.
 *
 * @param image - The DocxImage to convert.
 * @param options - Optional Typst converter options.
 * @returns Converted Typst image markup string.
 */
function convertImageToTypst(
  image: DocxImage,
  options?: TypstConverterOptions,
): string {
  const outputDir = options?.imageOutputDir ?? 'images';
  const cleanOutputDir = outputDir.replace(/[/\\]+$/, '');
  const filename = basename(image.zipPath.replace(/\\/g, '/'));
  const path = cleanOutputDir.length > 0 ? `${cleanOutputDir}/${filename}` : filename;

  if (image.widthEmu !== undefined) {
    const widthPct = Math.min(100, Math.round((image.widthEmu / 5486400) * 100));
    return `#image("${path}", width: ${widthPct}%)`;
  }

  return `#image("${path}")`;
}

/**
 * Joins converted runs preventing adjacent formatting markers from colliding.
 * When two adjacent runs both have formatting markers at their boundary
 * (such as `_`, `*`, `]`, `#`), a space is inserted to avoid ambiguous sequences like `**`.
 *
 * @param runs - Array of TextRun objects to convert and join.
 * @param escape - Whether to escape special characters.
 * @returns Joined Typst markup string.
 */
function joinConvertedRuns(runs: TextRun[], escape: boolean): string {
  const convertedRuns = runs.map((run) => convertRunToTypst(run, escape));
  let result = '';
  const isMarkerChar = (c: string) => c === '*' || c === '_' || c === ']' || c === '#';

  for (const current of convertedRuns) {
    if (result.length > 0 && current.length > 0) {
      const prevLast = result[result.length - 1];
      const currFirst = current[0];
      if (isMarkerChar(prevLast) && isMarkerChar(currFirst)) {
        result += ' ';
      }
    }
    result += current;
  }

  return result;
}

/**
 * Converts a measurement in twips (twentieths of a point) to centimeters formatted for Typst.
 * 1 twip = 1/1440 inch = 0.01764 cm (or 2.54 / 1440 cm).
 *
 * @param twips - Measurement in twips.
 * @returns Formatted measurement string with 'cm' unit (e.g. '1.27cm').
 */
export function twipsToCm(twips: number): string {
  return (twips / 1440 * 2.54).toFixed(2) + 'cm';
}

/**
 * Applies paragraph indentation (first-line, hanging, left, and right) to converted text content.
 * Does not apply to headings, list items, or empty content.
 *
 * @param content - Formatted Typst text content.
 * @param para - Source paragraph containing indent metadata.
 * @returns Content wrapped/prepended with Typst indentation markup.
 */
export function applyIndentation(content: string, para: Paragraph): string {
  if (
    !para.indent ||
    para.isEmpty ||
    para.listItem !== undefined ||
    getHeadingLevel(para.style) !== null ||
    content.length === 0
  ) {
    return content;
  }

  let result = content;

  // 1. First-line indent or Hanging indent (first-line position)
  if (para.indent.hanging !== undefined && para.indent.hanging > 0) {
    const yCm = twipsToCm(para.indent.hanging);
    result = `#h(-${yCm})${result}`;
  } else if (para.indent.firstLine !== undefined && para.indent.firstLine > 0) {
    const xCm = twipsToCm(para.indent.firstLine);
    result = `#h(${xCm})${result}`;
  }

  // 2. Left / Right indent (#pad)
  const padParts: string[] = [];
  if (para.indent.hanging !== undefined && para.indent.hanging > 0) {
    const leftTwips =
      para.indent.left !== undefined && para.indent.left > 0
        ? para.indent.left
        : para.indent.hanging;
    padParts.push(`left: ${twipsToCm(leftTwips)}`);
  } else if (para.indent.left !== undefined && para.indent.left > 0) {
    padParts.push(`left: ${twipsToCm(para.indent.left)}`);
  }

  if (para.indent.right !== undefined && para.indent.right > 0) {
    padParts.push(`right: ${twipsToCm(para.indent.right)}`);
  }

  if (padParts.length > 0) {
    result = `#pad(${padParts.join(', ')})[${result}]`;
  }

  return result;
}

/**
 * Converts a Paragraph into Typst markup.
 *
 * @param para - The paragraph to convert.
 * @param options - Converter options.
 * @returns Converted Typst markup for the paragraph, or an empty string for blank paragraphs.
 */
function convertParagraphToTypst(
  para: Paragraph,
  options?: TypstConverterOptions,
  isCell: boolean = false,
): string {
  const hasImages = Array.isArray(para.images) && para.images.length > 0;
  const hasPageBreakRun = para.runs?.some((r) => r.pageBreak);
  if (para.isEmpty && !hasImages && !hasPageBreakRun) {
    if (isCell) {
      return '';
    }
    let emptyOutput = '';
    if (para.pageBreakBefore) {
      emptyOutput = '#pagebreak()';
    }
    if (para.sectionBreak?.pageBreak) {
      emptyOutput = emptyOutput ? `${emptyOutput}\n#pagebreak()` : '#pagebreak()';
    }
    return emptyOutput;
  }

  const escape = options?.escapeSpecialChars !== false;
  let textContent = '';

  if (para.listItem !== undefined) {
    const runs =
      para.runs && para.runs.length > 0
        ? para.runs
        : para.text
          ? [{ text: para.text }]
          : [];

    const content = joinConvertedRuns(runs, escape);
    const indent = '  '.repeat(para.listItem.level);
    const bullet = para.listItem.listType === 'bullet' ? '-' : '+';
    textContent = `${indent}${bullet} ${content.trim()}`;
    // List items do not receive align wrapper
  } else {
    const headingLevel = getHeadingLevel(para.style);
    if (headingLevel !== null) {
      // Heading: unformatted text with level-specific '=' prefix
      const rawHeadingText =
        para.runs && para.runs.length > 0
          ? para.runs.map((r) => r.text).join('')
          : para.text || '';

      const text = escape ? escapeTypstText(rawHeadingText) : rawHeadingText;
      const trimmed = text.trim();
      const marker = '='.repeat(headingLevel);

      textContent = trimmed.length > 0 ? `${marker} ${trimmed}` : `${marker} `;
    } else {
      // Normal paragraph: format individual runs and handle whitespace
      const runs =
        para.runs && para.runs.length > 0
          ? para.runs
          : para.text
            ? [{ text: para.text }]
            : [];

      if (runs.length > 0) {
        const content = joinConvertedRuns(runs, escape);
        textContent = content.trim();
      }
    }

    // Apply indentation if specified (order: content -> #pad -> #align)
    textContent = applyIndentation(textContent, para);

    // Apply alignment if specified
    if (textContent.length > 0) {
      if (para.align === 'center') {
        textContent = `#align(center)[${textContent}]`;
      } else if (para.align === 'right') {
        textContent = `#align(right)[${textContent}]`;
      }
    }
  }

  let output = textContent;
  if (hasImages) {
    const imageMarkups = para.images!.map((img) => convertImageToTypst(img, options));
    if (output.length > 0) {
      output = [output, ...imageMarkups].join('\n');
    } else {
      output = imageMarkups.join('\n');
    }
  }

  if (output.length > 0 && !isCell) {
    if (para.hasBorderTop) {
      output = `#line(length: 100%)\n${output}`;
    }
    if (para.hasBorderBottom) {
      output = `${output}\n#line(length: 100%)`;
    }
  }

  if (!isCell) {
    if (para.pageBreakBefore) {
      output = output.length > 0 ? `#pagebreak()\n${output}` : '#pagebreak()';
    }
    if (para.sectionBreak?.pageBreak) {
      output = output.length > 0 ? `${output}\n#pagebreak()` : '#pagebreak()';
    }
  }

  return output;
}

/**
 * Converts a TableCell into Typst content wrapped in square brackets `[...]`.
 *
 * Joins all paragraphs within the cell, converting individual runs with formatting,
 * and separating paragraphs with newline characters. Empty cells produce `[]`.
 *
 * @param cell - The table cell to convert.
 * @param options - Optional Typst converter configuration.
 * @returns Converted Typst cell markup string.
 */
function convertCellToTypst(cell: TableCell, options?: TypstConverterOptions): string {
  if (cell.vMerge === 'continue' || (cell.isVerticalMerge && cell.vMerge === undefined && !cell.rowSpan)) {
    return '[]';
  }

  const paraTexts: string[] = [];

  if (cell.paragraphs && cell.paragraphs.length > 0) {
    for (const para of cell.paragraphs) {
      if (para.isEmpty && (!para.images || para.images.length === 0)) {
        continue;
      }
      const converted = convertParagraphToTypst(para, options, true);
      if (converted.length > 0) {
        paraTexts.push(converted);
      }
    }
  }

  let content = paraTexts.length > 0 ? paraTexts.join('\n\n') : '';
  if (content.trim().length > 0) {
    content = `#set par(spacing: 0.5em); ${content}`;
  }

  const attrs: string[] = [];
  if (cell.columnSpan && cell.columnSpan > 1) {
    attrs.push(`colspan: ${cell.columnSpan}`);
  }
  if (cell.rowSpan && cell.rowSpan > 1) {
    attrs.push(`rowspan: ${cell.rowSpan}`);
  }
  if (cell.align) {
    attrs.push(`align: ${cell.align}`);
  }

  if (attrs.length > 0) {
    return `table.cell(${attrs.join(', ')})[${content}]`;
  }

  return `[${content}]`;
}

/**
 * Converts a structured DocxTable into Typst `#table(...)` markup.
 *
 * Formats table columns, emits `table.header(...)` blocks for header rows,
 * and formats all cell content with indentation and comma delimiters.
 *
 * @param table - The table structure to convert.
 * @param options - Optional Typst converter options.
 * @returns Converted Typst table markup string.
 *
 * @example
 * ```typescript
 * const typstTable = convertTableToTypst(table);
 * console.log(typstTable);
 * ```
 */
export function convertTableToTypst(
  table: DocxTable,
  options?: TypstConverterOptions,
): string {
  let columnsDecl = `  columns: ${table.columnCount},`;
  if (
    Array.isArray(table.columnWidths) &&
    table.columnWidths.length === table.columnCount
  ) {
    const totalWidth = table.columnWidths.reduce((a, b) => a + b, 0);
    if (totalWidth > 0) {
      const fractions = table.columnWidths.map((w) => Math.round((w / totalWidth) * 100));
      columnsDecl = `  columns: (${fractions.map((f) => `${f}%`).join(', ')}),`;
    }
  }

  const lines: string[] = [
    '#table(',
    columnsDecl,
  ];

  const borders = table.borders;
  const isHorizontalOnly =
    borders !== undefined &&
    (borders.top || borders.bottom) &&
    !borders.left &&
    !borders.right &&
    !borders.insideV;

  if (isHorizontalOnly) {
    lines.push('  stroke: none,');
    if (borders.top !== false) {
      lines.push('  table.hline(y: 0),');
    }
  } else if (!borders) {
    lines.push('  stroke: 1pt,');
  } else {
    const hasV = borders.left || borders.right || borders.insideV;
    const hasH = borders.top || borders.bottom || borders.insideH;
    if (hasV && hasH) {
      lines.push('  stroke: 1pt,');
    } else if (hasV && !hasH) {
      lines.push('  stroke: (x: 1pt, y: none),');
    } else {
      lines.push('  stroke: 1pt,');
    }
  }

  for (const row of table.rows) {
    let slotsUsed = 0;
    const rowCells: string[] = [];
    for (const cell of row.cells) {
      if (slotsUsed >= table.columnCount) {
        break; // skip trailing overflow cells
      }
      if (cell.vMerge === 'continue') {
        slotsUsed += cell.columnSpan ?? 1;
        continue;
      }
      rowCells.push(convertCellToTypst(cell, options));
      slotsUsed += cell.columnSpan ?? 1;
    }
    if (rowCells.length === 0) {
      continue;
    }
    const cellsContent = rowCells.join(', ');

    if (row.isHeader) {
      lines.push('  table.header(');
      lines.push(`    ${cellsContent},`);
      lines.push('  ),');
    } else {
      lines.push(`  ${cellsContent},`);
    }
  }

  if (isHorizontalOnly && borders?.bottom !== false) {
    lines.push('  table.hline(),');
  }

  lines.push(')');
  let result = lines.join('\n');
  if (table.hasBorderTop && (!borders || borders.top !== true)) {
    result = `#line(length: 100%)\n${result}`;
  }
  if (table.hasBorderBottom && (!borders || borders.bottom !== true)) {
    result = `${result}\n#line(length: 100%)`;
  }
  return result;
}

/**
 * Converts a parsed DocxDocument into a TypstDocument containing Typst markup
 * and conversion statistics.
 *
 * @param doc - The parsed DOCX document structure.
 * @param options - Optional configuration options for the Typst conversion.
 * @returns The converted TypstDocument including markup content and statistics.
 *
 * @example
 * ```typescript
 * import { convertDocxToTypst } from 'docx2typst';
 *
 * const typstDoc = convertDocxToTypst(docxDocument, {
 *   includeHeader: true,
 *   paragraphSpacing: true,
 * });
 * console.log(typstDoc.content);
 * ```
 */
export function convertDocxToTypst(
  doc: DocxDocument,
  options?: TypstConverterOptions,
): TypstDocument {
  const paragraphSpacing = options?.paragraphSpacing !== false;

  let paragraphCount = 0;
  let headingCount = 0;
  let runsWithFormatting = 0;

  for (const para of doc.paragraphs) {
    const headingLevel = para.listItem ? null : getHeadingLevel(para.style);
    if (headingLevel !== null) {
      headingCount++;
    } else {
      paragraphCount++;
    }

    const runs = para.runs || [];
    for (const run of runs) {
      if (run.bold || run.italic || run.underline || run.color) {
        runsWithFormatting++;
      }
    }
  }

  const items: BodyItem[] =
    doc.bodyItems && doc.bodyItems.length > 0
      ? doc.bodyItems
      : [
          ...doc.paragraphs.map((p) => ({ type: 'paragraph' as const, paragraph: p })),
          ...(doc.tables || []).map((t) => ({ type: 'table' as const, table: t })),
        ];

  interface SectionGroup {
    items: BodyItem[];
    columnCount: number;
  }

  const sections: SectionGroup[] = [];
  let currentGroupItems: BodyItem[] = [];

  for (const item of items) {
    currentGroupItems.push(item);
    if (item.type === 'paragraph' && item.paragraph.sectionBreak) {
      sections.push({
        items: currentGroupItems,
        columnCount: item.paragraph.sectionBreak.columnCount,
      });
      currentGroupItems = [];
    }
  }

  if (currentGroupItems.length > 0 || sections.length === 0) {
    const finalColumnCount =
      doc.sections && doc.sections.length > 0
        ? doc.sections[doc.sections.length - 1].columnCount
        : 1;
    sections.push({
      items: currentGroupItems,
      columnCount: finalColumnCount,
    });
  }

  type RenderedItemEntry = BodyItem | { type: 'showRule' };
  const renderedParts: string[] = [];
  const filteredItems: RenderedItemEntry[] = [];
  let consecutiveBlankCount = 0;
  let showColumnsEmitted = false;

  const itemColCounts: number[] = [];
  for (const sec of sections) {
    for (let k = 0; k < sec.items.length; k++) {
      itemColCounts.push(sec.columnCount);
    }
  }

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const colCount = itemColCounts[i] ?? 1;

    if (item.type === 'paragraph') {
      const hasImages = Array.isArray(item.paragraph.images) && item.paragraph.images.length > 0;
      const isHeading = !item.paragraph.listItem && getHeadingLevel(item.paragraph.style) !== null;
      const hasPageBreak = Boolean(
        item.paragraph.pageBreakBefore ||
          item.paragraph.sectionBreak?.pageBreak ||
          item.paragraph.runs?.some((r) => r.pageBreak),
      );
      const isBlank =
        !isHeading &&
        !hasPageBreak &&
        (Boolean(item.paragraph.isEmpty) ||
          ((!item.paragraph.text || item.paragraph.text.trim().length === 0) && !hasImages));

      if (isBlank) {
        consecutiveBlankCount++;
        if (consecutiveBlankCount > 1) {
          continue;
        }
        renderedParts.push('');
        filteredItems.push(item);
      } else {
        consecutiveBlankCount = 0;
        if (!showColumnsEmitted && colCount > 1) {
          renderedParts.push(`#show: rest => columns(${colCount}, rest)`);
          filteredItems.push({ type: 'showRule' });
          showColumnsEmitted = true;
        }
        renderedParts.push(convertParagraphToTypst(item.paragraph, options));
        filteredItems.push(item);
      }
    } else {
      consecutiveBlankCount = 0;
      if (!showColumnsEmitted && colCount > 1) {
        renderedParts.push(`#show: rest => columns(${colCount}, rest)`);
        filteredItems.push({ type: 'showRule' });
        showColumnsEmitted = true;
      }
      renderedParts.push(convertTableToTypst(item.table, options));
      filteredItems.push(item);
    }
  }

  let content = '';
  for (let i = 0; i < renderedParts.length; i++) {
    if (i > 0) {
      if (!paragraphSpacing) {
        content += '\n';
      } else {
        const prevItem = filteredItems[i - 1];
        const currItem = filteredItems[i];
        const isPrevList =
          prevItem.type === 'paragraph' && prevItem.paragraph.listItem !== undefined;
        const isCurrList =
          currItem.type === 'paragraph' && currItem.paragraph.listItem !== undefined;
        if (isPrevList && isCurrList) {
          content += '\n';
        } else {
          content += '\n\n';
        }
      }
    }
    content += renderedParts[i];
  }

  if (options?.includeHeader) {
    content = `// Generated by docx2typst\n\n${content}`;
  }

  return {
    content,
    stats: {
      paragraphCount,
      headingCount,
      runsWithFormatting,
    },
  };
}
