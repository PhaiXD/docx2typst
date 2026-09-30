/**
 * @file types.ts
 * Type definitions and data interfaces for the docx2typst library.
 */

/**
 * Configuration options for DOCX parsing and text extraction.
 */
export interface DocxParserOptions {
  /**
   * Whether to extract and include paragraph style identifiers.
   * @default true
   */
  includeStyles?: boolean;

  /**
   * Whether to preserve whitespace when specified by xml:space="preserve".
   * @default true
   */
  preserveWhitespace?: boolean;
}

/**
 * Raw XML content extracted from the underlying .docx ZIP archive.
 */
export interface RawXmlContent {
  /**
   * The raw XML string of the main document part (`word/document.xml`).
   */
  documentXml: string;

  /**
   * The raw XML string of the document relationships part (`word/_rels/document.xml.rels`), if present.
   */
  relationshipsXml?: string;

  /**
   * The raw XML string of the numbering definitions part (`word/numbering.xml`), if present.
   */
  numberingXml?: string;
}

/**
 * Represents a styled run of text (`w:r`) within a paragraph.
 */
export interface TextRun {
  /**
   * The text content of the run (`w:t`).
   */
  text: string;

  /**
   * Whether the run is formatted as bold (`w:b`).
   */
  bold?: boolean;

  /**
   * Whether the run is formatted as italic (`w:i`).
   */
  italic?: boolean;

  /**
   * Whether the run is formatted as underlined (`w:u`).
   */
  underline?: boolean;
}

/**
 * Information describing a list item paragraph in an OOXML document.
 */
export interface ListItemInfo {
  /**
   * The abstract numbering definition ID.
   */
  abstractNumId: number;

  /**
   * The numbering instance ID referenced by the paragraph.
   */
  numId: number;

  /**
   * The nesting / indentation level (0-based).
   */
  level: number;

  /**
   * The resolved list type: unordered bullet or ordered numbered list.
   */
  listType: 'bullet' | 'ordered';
}

/**
 * Mapping from numbering instance ID (`w:numId`) to abstract numbering ID (`w:abstractNumId`).
 */
export type NumIdMap = Map<number, number>;

/**
 * Mapping from compound key `"${abstractNumId}:${level}"` to list type (`bullet` or `ordered`).
 */
export type AbstractNumMap = Map<string, 'bullet' | 'ordered'>;

/**
 * Represents a paragraph (`w:p`) within a Word document body.
 */
export interface Paragraph {
  /**
   * The combined plain text content of all runs in this paragraph.
   */
  text: string;

  /**
   * The individual text runs contained within this paragraph.
   */
  runs: TextRun[];

  /**
   * The style identifier of the paragraph (`w:pPr > w:pStyle @w:val`), if present.
   * Examples include 'Heading1', 'Heading2', 'Normal'.
   */
  style?: string;

  /**
   * Indicates whether the paragraph has no text content.
   */
  isEmpty?: boolean;

  /**
   * List item numbering and nesting metadata, if this paragraph is part of a list.
   */
  listItem?: ListItemInfo;
}

/**
 * Represents the structured content of a parsed .docx document.
 */
export interface DocxDocument {
  /**
   * The list of paragraphs extracted from the document body in order of appearance.
   */
  paragraphs: Paragraph[];

  /**
   * The complete plain text of the entire document, with paragraphs joined by newlines.
   */
  text: string;
}

/**
 * Options for Typst output generation.
 */
export interface TypstConverterOptions {
  /**
   * Whether to add a blank line between paragraphs.
   * @default true
   */
  paragraphSpacing?: boolean;

  /**
   * Whether to escape special Typst characters in text.
   * @default true
   */
  escapeSpecialChars?: boolean;

  /**
   * Whether to include a Typst document header comment.
   * @default false
   */
  includeHeader?: boolean;
}

/**
 * The result of converting a DocxDocument to Typst format.
 */
export interface TypstDocument {
  /**
   * The complete Typst markup string.
   */
  content: string;

  /**
   * Statistics about the conversion.
   */
  stats: {
    paragraphCount: number;
    headingCount: number;
    runsWithFormatting: number;
  };
}

