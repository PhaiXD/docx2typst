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
  if (!run.text || run.text.length === 0) {
    return '';
  }

  const text = escape ? escapeTypstText(run.text) : run.text;

  // If no formatting is active, return the text directly
  if (!run.bold && !run.italic && !run.underline) {
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
 * Converts a Paragraph into Typst markup.
 *
 * @param para - The paragraph to convert.
 * @param options - Converter options.
 * @returns Converted Typst markup for the paragraph, or an empty string for blank paragraphs.
 */
function convertParagraphToTypst(
  para: Paragraph,
  options?: TypstConverterOptions,
): string {
  const hasImages = Array.isArray(para.images) && para.images.length > 0;
  if (para.isEmpty && !hasImages) {
    return '';
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

    const content = runs.map((run) => convertRunToTypst(run, escape)).join('');
    const indent = '  '.repeat(para.listItem.level);
    const bullet = para.listItem.listType === 'bullet' ? '-' : '+';
    textContent = `${indent}${bullet} ${content.trim()}`;
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
        const content = runs.map((run) => convertRunToTypst(run, escape)).join('');
        textContent = content.trim();
      }
    }
  }

  if (hasImages) {
    const imageMarkups = para.images!.map((img) => convertImageToTypst(img, options));
    if (textContent.length > 0) {
      return [textContent, ...imageMarkups].join('\n');
    }
    return imageMarkups.join('\n');
  }

  return textContent;
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
  if (cell.isVerticalMerge) {
    return '[]';
  }

  const escape = options?.escapeSpecialChars !== false;
  const paraTexts: string[] = [];

  if (cell.paragraphs && cell.paragraphs.length > 0) {
    for (const para of cell.paragraphs) {
      if (para.isEmpty) {
        continue;
      }
      const runs =
        para.runs && para.runs.length > 0
          ? para.runs
          : para.text
            ? [{ text: para.text }]
            : [];

      const paraContent = runs.map((run) => convertRunToTypst(run, escape)).join('').trim();
      if (para.images && para.images.length > 0) {
        const imageMarkups = para.images.map((img) => convertImageToTypst(img, options));
        if (paraContent.length > 0) {
          paraTexts.push([paraContent, ...imageMarkups].join('\n'));
        } else {
          paraTexts.push(imageMarkups.join('\n'));
        }
      } else if (paraContent.length > 0) {
        paraTexts.push(paraContent);
      }
    }
  }

  const content = paraTexts.length > 0 ? paraTexts.join('\n') : '';

  const attrs: string[] = [];
  if (cell.columnSpan && cell.columnSpan > 1) {
    attrs.push(`colspan: ${cell.columnSpan}`);
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
  const lines: string[] = [
    '#table(',
    `  columns: ${table.columnCount},`,
    '  stroke: 1pt,',
  ];

  for (const row of table.rows) {
    const cellsContent = row.cells
      .map((cell) => convertCellToTypst(cell, options))
      .join(', ');

    if (row.isHeader) {
      lines.push('  table.header(');
      lines.push(`    ${cellsContent},`);
      lines.push('  ),');
    } else {
      lines.push(`  ${cellsContent},`);
    }
  }

  lines.push(')');
  return lines.join('\n');
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
      if (run.bold || run.italic || run.underline) {
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

  function renderSectionItems(sectionItems: BodyItem[]): string {
    const renderedParts: string[] = [];

    for (const item of sectionItems) {
      if (item.type === 'paragraph') {
        renderedParts.push(convertParagraphToTypst(item.paragraph, options));
      } else {
        renderedParts.push(convertTableToTypst(item.table, options));
      }
    }

    let sectionContent = '';
    for (let i = 0; i < renderedParts.length; i++) {
      if (i > 0) {
        if (!paragraphSpacing) {
          sectionContent += '\n';
        } else {
          const prevItem = sectionItems[i - 1];
          const currItem = sectionItems[i];
          const isPrevList =
            prevItem.type === 'paragraph' && prevItem.paragraph.listItem !== undefined;
          const isCurrList =
            currItem.type === 'paragraph' && currItem.paragraph.listItem !== undefined;
          if (isPrevList && isCurrList) {
            sectionContent += '\n';
          } else {
            sectionContent += '\n\n';
          }
        }
      }
      sectionContent += renderedParts[i];
    }

    return sectionContent;
  }

  const renderedSections: string[] = [];

  for (const section of sections) {
    const rawContent = renderSectionItems(section.items);
    if (rawContent.length === 0) {
      continue;
    }

    if (section.columnCount > 1) {
      const indented = rawContent
        .split('\n')
        .map((line) => (line.length > 0 ? `  ${line}` : ''))
        .join('\n');
      renderedSections.push(`#columns(${section.columnCount})[\n${indented}\n]`);
    } else {
      renderedSections.push(rawContent);
    }
  }

  const separator = paragraphSpacing ? '\n\n' : '\n';
  let content = renderedSections.join(separator);

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
