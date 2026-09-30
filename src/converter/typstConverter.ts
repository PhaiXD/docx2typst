/**
 * @file typstConverter.ts
 * Converts structured DocxDocument content into Typst markup format.
 */

import type {
  DocxDocument,
  Paragraph,
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
  if (para.isEmpty) {
    return '';
  }

  const escape = options?.escapeSpecialChars !== false;

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
    return `${indent}${bullet} ${content.trim()}`;
  }

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

    return trimmed.length > 0 ? `${marker} ${trimmed}` : `${marker} `;
  }

  // Normal paragraph: format individual runs and handle whitespace
  const runs =
    para.runs && para.runs.length > 0
      ? para.runs
      : para.text
        ? [{ text: para.text }]
        : [];

  if (runs.length === 0) {
    return '';
  }

  const content = runs.map((run) => convertRunToTypst(run, escape)).join('');
  return content.trim();
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

  const convertedParagraphs: string[] = [];

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

    convertedParagraphs.push(convertParagraphToTypst(para, options));
  }

  let content = '';
  for (let i = 0; i < convertedParagraphs.length; i++) {
    if (i > 0) {
      if (!paragraphSpacing) {
        content += '\n';
      } else {
        const prevPara = doc.paragraphs[i - 1];
        const currPara = doc.paragraphs[i];
        const isPrevList = prevPara.listItem !== undefined;
        const isCurrList = currPara.listItem !== undefined;
        if (isPrevList && isCurrList) {
          content += '\n';
        } else {
          content += '\n\n';
        }
      }
    }
    content += convertedParagraphs[i];
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
