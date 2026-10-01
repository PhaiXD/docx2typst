/**
 * @file types.ts
 * Type definitions and data interfaces for the docx2typst library.
 */

/**
 * Options controlling how images are extracted from the DOCX archive.
 */
export interface ImageExtractionOptions {
  /**
   * The output directory where extracted images should be saved on disk.
   * @default 'images'
   */
  outputDir?: string;

  /**
   * Whether image extraction is enabled.
   * @default true
   */
  enabled?: boolean;
}

/**
 * Represents an embedded image extracted from an OOXML Word document.
 */
export interface DocxImage {
  /**
   * The relationship ID referencing this image in the document relationships part (`word/_rels/document.xml.rels`).
   */
  relationshipId: string;

  /**
   * The relative target path from relationships (e.g. `'media/image1.png'`).
   */
  targetPath: string;

  /**
   * The full entry path within the .docx ZIP archive (e.g. `'word/media/image1.png'`).
   */
  zipPath: string;

  /**
   * The detected MIME type of the image (e.g. `'image/png'`).
   */
  mimeType?: string;

  /**
   * Descriptive alternative text extracted from `wp:docPr @descr` or `@title`.
   */
  altText?: string;

  /**
   * Width of the image in English Metric Units (EMUs), where 1 inch = 914,400 EMUs.
   */
  widthEmu?: number;

  /**
   * Height of the image in English Metric Units (EMUs), where 1 inch = 914,400 EMUs.
   */
  heightEmu?: number;
}

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

  /**
   * Image extraction configuration options.
   */
  images?: ImageExtractionOptions;

  /**
   * Raw XML content of word/styles.xml for style inheritance resolution.
   */
  stylesXml?: string;
}

/**
 * Paragraph indentation properties in twips (1 inch = 1440 twips).
 */
export interface ParagraphIndent {
  /**
   * Left indentation in twips.
   */
  left?: number;

  /**
   * Right indentation in twips.
   */
  right?: number;

  /**
   * First line indentation in twips.
   */
  firstLine?: number;

  /**
   * Hanging indentation in twips.
   */
  hanging?: number;
}

/**
 * Metadata for a paragraph style extracted from word/styles.xml.
 */
export interface DocxStyleInfo {
  /**
   * The style identifier (e.g. 'Normal', 'Title', 'Heading1', 'Paper-Title').
   */
  styleId: string;

  /**
   * The user-friendly style name (from w:name @w:val), if present.
   */
  name?: string;

  /**
   * The base style identifier this style inherits from (from w:basedOn @w:val).
   */
  basedOn?: string;

  /**
   * The paragraph alignment defined by this style.
   */
  align?: 'left' | 'center' | 'right' | 'justify';

  /**
   * Paragraph indentation defined by this style.
   */
  indent?: ParagraphIndent;

  /**
   * Paragraph spacing defined by this style in twips.
   */
  spacing?: {
    after?: number;
    before?: number;
    line?: number;
  };

  /**
   * Whether this is a default style.
   */
  isDefault?: boolean;
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

  /**
   * The raw XML string of the styles part (`word/styles.xml`), if present.
   */
  stylesXml?: string;

  /**
   * Map of relationship IDs to extracted DocxImage metadata.
   */
  imageMap?: Map<string, DocxImage>;

  /**
   * The raw JSZip archive instance, stored as unknown to decouple types from JSZip.
   */
  zipInstance?: unknown;
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

  /**
   * The text color in hex format (e.g. 'ff0000', '0000ff').
   */
  color?: string;

  /**
   * Whether this run contains an inline page break (`w:br w:type="page"`).
   */
  pageBreak?: boolean;
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

  /**
   * Numbering format string (e.g. 'bullet', 'decimal', 'lowerLetter', 'lowerRoman').
   */
  numFmt?: string;
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
 * Mapping from compound key `"${abstractNumId}:${level}"` to numbering format string (`w:numFmt @w:val`).
 */
export type AbstractNumFmtMap = Map<string, string>;

/**
 * Section properties describing page or multi-column layout.
 */
export interface SectionProperties {
  columnCount: number;

  /**
   * Whether this section break creates a new page.
   */
  pageBreak?: boolean;
}

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
   * Whether this paragraph forces a page break before it (`w:pPr > w:pageBreakBefore`).
   */
  pageBreakBefore?: boolean;

  /**
   * List item numbering and nesting metadata, if this paragraph is part of a list.
   */
  listItem?: ListItemInfo;

  /**
   * Embedded images contained within this paragraph.
   */
  images?: DocxImage[];

  /**
   * Section break properties defined at the end of this paragraph.
   */
  sectionBreak?: SectionProperties;

  /**
   * Paragraph text alignment ('left' | 'center' | 'right' | 'justify').
   */
  align?: 'left' | 'center' | 'right' | 'justify';

  /**
   * Whether the paragraph has a top border (from w:pBdr > w:top).
   */
  hasBorderTop?: boolean;

  /**
   * Whether the paragraph has a bottom border (from w:pBdr > w:bottom).
   */
  hasBorderBottom?: boolean;

  /**
   * Indentation properties of the paragraph in twips.
   */
  indent?: ParagraphIndent;

  /**
   * Paragraph spacing properties in twips.
   */
  spacing?: {
    after?: number;    // in twips
    before?: number;   // in twips
    line?: number;     // in twips
  };
}

/**
 * Represents a single cell (`w:tc`) within a table row.
 */
export interface TableCell {
  /**
   * The list of paragraphs contained within this cell.
   */
  paragraphs: Paragraph[];

  /**
   * The number of grid columns spanned by this cell (`w:gridSpan`), if greater than 1.
   */
  columnSpan?: number;

  /**
   * The number of table rows spanned by this cell (`w:vMerge`), if greater than 1.
   */
  rowSpan?: number;

  /**
   * Indicates whether this cell is vertically merged (`w:vMerge`).
   */
  isVerticalMerge?: boolean;

  /**
   * Vertical merge state: 'restart' starts a merge, 'continue' continues a merge from above.
   */
  vMerge?: 'restart' | 'continue';

  /**
   * Cell vertical alignment in Typst ('top' | 'horizon' | 'bottom' | 'left' | 'center' | 'right').
   */
  align?: 'top' | 'horizon' | 'bottom' | 'left' | 'center' | 'right' | string;
}

/**
 * Represents a single row (`w:tr`) within a table.
 */
export interface TableRow {
  /**
   * The cells contained within this table row.
   */
  cells: TableCell[];

  /**
   * Indicates whether this row is a header row.
   */
  isHeader?: boolean;
}

/**
 * Represents a structured table (`w:tbl`) within a Word document body.
 */
export interface DocxTable {
  /**
   * The rows contained within this table.
   */
  rows: TableRow[];

  /**
   * Total number of columns in the table grid.
   */
  columnCount: number;

  /**
   * The style identifier of the table (`w:tblPr > w:tblStyle @w:val`), if present.
   */
  style?: string;

  /**
   * Column widths in twips from w:tblGrid > w:gridCol.
   */
  columnWidths?: number[];

  /**
   * Table border configuration from w:tblPr > w:tblBorders or cell overrides.
   */
  borders?: {
    top?: boolean;
    bottom?: boolean;
    left?: boolean;
    right?: boolean;
    insideH?: boolean;
    insideV?: boolean;
  };

  /**
   * Whether the table has a top border (e.g. from w:tblBorders, cell borders, or first row paragraph borders).
   */
  hasBorderTop?: boolean;

  /**
   * Whether the table has a bottom border (e.g. from w:tblBorders, cell borders, or last row paragraph borders).
   */
  hasBorderBottom?: boolean;
}

/**
 * Represents an item in document order within the document body.
 */
export type BodyItem =
  | { type: 'paragraph'; paragraph: Paragraph }
  | { type: 'table'; table: DocxTable };

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

  /**
   * The list of tables extracted from the document body in order of appearance.
   */
  tables?: DocxTable[];

  /**
   * Document-level section properties (e.g. final section multi-column layout).
   */
  sections?: SectionProperties[];

  /**
   * Document items in appearance order, interleaving paragraphs and tables.
   */
  bodyItems?: BodyItem[];
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

  /**
   * Output directory for images in Typst `#image(...)` paths.
   * @default 'images'
   */
  imageOutputDir?: string;
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

