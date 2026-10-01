/**
 * @file typstConverter.test.ts
 * Unit tests for OOXML to Typst markup conversion layer.
 */

import {
  convertDocxToTypst,
  convertTableToTypst,
  getHeadingLevel,
  twipsToCm,
} from '../converter/typstConverter.js';
import type { DocxDocument, DocxTable, Paragraph } from '../types.js';

describe('typstConverter', () => {
  describe('getHeadingLevel', () => {
    it('should map standard OOXML heading styles (Heading1 - Heading6)', () => {
      expect(getHeadingLevel('Heading1')).toBe(1);
      expect(getHeadingLevel('Heading2')).toBe(2);
      expect(getHeadingLevel('Heading3')).toBe(3);
      expect(getHeadingLevel('Heading4')).toBe(4);
      expect(getHeadingLevel('Heading5')).toBe(5);
      expect(getHeadingLevel('Heading6')).toBe(6);
    });

    it('should map lowercase heading styles', () => {
      expect(getHeadingLevel('heading1')).toBe(1);
      expect(getHeadingLevel('heading2')).toBe(2);
      expect(getHeadingLevel('heading6')).toBe(6);
    });

    it('should map single digit heading styles', () => {
      expect(getHeadingLevel('1')).toBe(1);
      expect(getHeadingLevel('2')).toBe(2);
      expect(getHeadingLevel('6')).toBe(6);
    });

    it('should map heading styles with space', () => {
      expect(getHeadingLevel('Heading 1')).toBe(1);
      expect(getHeadingLevel('heading 2')).toBe(2);
      expect(getHeadingLevel('HEADING 6')).toBe(6);
    });

    it('should return null for non-heading styles and levels beyond 6', () => {
      expect(getHeadingLevel('Heading7')).toBeNull();
      expect(getHeadingLevel('Heading 7')).toBeNull();
      expect(getHeadingLevel('7')).toBeNull();
      expect(getHeadingLevel('0')).toBeNull();
      expect(getHeadingLevel('Normal')).toBeNull();
      expect(getHeadingLevel('Title')).toBeNull();
      expect(getHeadingLevel('Subtitle')).toBeNull();
      expect(getHeadingLevel('')).toBeNull();
      expect(getHeadingLevel('   ')).toBeNull();
      expect(getHeadingLevel(undefined)).toBeNull();
    });
  });

  describe('convertDocxToTypst - text formatting', () => {
    it('should convert a plain paragraph without formatting', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Hello world',
            runs: [{ text: 'Hello world' }],
          },
        ],
        text: 'Hello world',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Hello world');
      expect(result.stats.paragraphCount).toBe(1);
      expect(result.stats.headingCount).toBe(0);
      expect(result.stats.runsWithFormatting).toBe(0);
    });

    it('should convert bold text with asterisks', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Bold text',
            runs: [{ text: 'Bold text', bold: true }],
          },
        ],
        text: 'Bold text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('*Bold text*');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert italic text with underscores', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Italic text',
            runs: [{ text: 'Italic text', italic: true }],
          },
        ],
        text: 'Italic text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('_Italic text_');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert underline text with #underline[]', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Underlined text',
            runs: [{ text: 'Underlined text', underline: true }],
          },
        ],
        text: 'Underlined text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#underline[Underlined text]');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert bold and italic combined with *_text_* nesting', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Bold Italic text',
            runs: [{ text: 'Bold Italic text', bold: true, italic: true }],
          },
        ],
        text: 'Bold Italic text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('*_Bold Italic text_*');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert bold and underline combined with #underline[*text*]', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Bold Underlined',
            runs: [{ text: 'Bold Underlined', bold: true, underline: true }],
          },
        ],
        text: 'Bold Underlined',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#underline[*Bold Underlined*]');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert all three (bold, italic, underline) with correct nesting order', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'All styled',
            runs: [{ text: 'All styled', bold: true, italic: true, underline: true }],
          },
        ],
        text: 'All styled',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#underline[*_All styled_*]');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should convert multiple runs within a single paragraph correctly', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'This is bold and italic text.',
            runs: [
              { text: 'This is ' },
              { text: 'bold', bold: true },
              { text: ' and ' },
              { text: 'italic', italic: true },
              { text: ' text.' },
            ],
          },
        ],
        text: 'This is bold and italic text.',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('This is *bold* and _italic_ text.');
      expect(result.stats.runsWithFormatting).toBe(2);
    });

    it('should handle runs with leading and trailing whitespace properly', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'start bold end',
            runs: [
              { text: 'start' },
              { text: ' bold ', bold: true },
              { text: 'end' },
            ],
          },
        ],
        text: 'start bold end',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('start *bold* end');
    });

    it('should handle whitespace-only formatted runs gracefully', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'a   b',
            runs: [
              { text: 'a' },
              { text: '   ', bold: true },
              { text: 'b' },
            ],
          },
        ],
        text: 'a   b',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('a   b');
    });

    it('should handle empty text runs gracefully', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Non-empty',
            runs: [
              { text: '', bold: true },
              { text: 'Non-empty' },
              { text: '', italic: true },
            ],
          },
        ],
        text: 'Non-empty',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Non-empty');
    });
  });

  describe('convertDocxToTypst - headings', () => {
    it('should convert Heading 1 with "= " prefix', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Main Title',
            runs: [{ text: 'Main Title' }],
            style: 'Heading1',
          },
        ],
        text: 'Main Title',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('= Main Title');
      expect(result.stats.headingCount).toBe(1);
      expect(result.stats.paragraphCount).toBe(0);
    });

    it('should convert Heading 2 with "== " prefix', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Section Title',
            runs: [{ text: 'Section Title' }],
            style: 'Heading2',
          },
        ],
        text: 'Section Title',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('== Section Title');
      expect(result.stats.headingCount).toBe(1);
    });

    it('should convert Heading 6 with "====== " prefix', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Deep Section',
            runs: [{ text: 'Deep Section' }],
            style: 'Heading6',
          },
        ],
        text: 'Deep Section',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('====== Deep Section');
      expect(result.stats.headingCount).toBe(1);
    });

    it('should strip run formatting inside headings', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Formatted Heading',
            runs: [
              { text: 'Formatted ', bold: true },
              { text: 'Heading', italic: true },
            ],
            style: 'Heading 1',
          },
        ],
        text: 'Formatted Heading',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('= Formatted Heading');
    });

    it('should handle empty headings by providing heading marker prefix', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            style: 'Heading1',
            isEmpty: false,
          },
          {
            text: '',
            runs: [],
            style: 'Heading2',
            isEmpty: false,
          },
          {
            text: '',
            runs: [],
            style: 'Heading6',
            isEmpty: false,
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('= \n\n== \n\n======');
    });
  });

  describe('convertDocxToTypst - special character escaping', () => {
    it('should escape @, #, <, >, `, and \\ in text by default', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Special: @user #channel <tag> `code` back\\slash',
            runs: [{ text: 'Special: @user #channel <tag> `code` back\\slash' }],
          },
        ],
        text: 'Special: @user #channel <tag> `code` back\\slash',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe(
        'Special: \\@user \\#channel \\<tag\\> \\`code\\` back\\\\slash',
      );
    });

    it('should escape special characters inside formatted runs without escaping formatting markers', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'tag #1 @admin',
            runs: [{ text: 'tag #1 @admin', bold: true }],
          },
        ],
        text: 'tag #1 @admin',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('*tag \\#1 \\@admin*');
    });

    it('should escape special characters in headings', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'C# & C++ <Guide>',
            runs: [{ text: 'C# & C++ <Guide>' }],
            style: 'Heading1',
          },
        ],
        text: 'C# & C++ <Guide>',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('= C\\# & C++ \\<Guide\\>');
    });

    it('should not escape special characters when escapeSpecialChars is false', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Special: @user #channel <tag> `code` back\\slash',
            runs: [{ text: 'Special: @user #channel <tag> `code` back\\slash' }],
          },
        ],
        text: 'Special: @user #channel <tag> `code` back\\slash',
      };

      const result = convertDocxToTypst(doc, { escapeSpecialChars: false });
      expect(result.content).toBe(
        'Special: @user #channel <tag> `code` back\\slash',
      );
    });
  });

  describe('convertDocxToTypst - paragraph spacing and empty paragraphs', () => {
    it('should separate multiple paragraphs with blank line (\\n\\n) by default', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'First paragraph',
            runs: [{ text: 'First paragraph' }],
          },
          {
            text: 'Second paragraph',
            runs: [{ text: 'Second paragraph' }],
          },
        ],
        text: 'First paragraph\nSecond paragraph',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('First paragraph\n\nSecond paragraph');
      expect(result.stats.paragraphCount).toBe(2);
    });

    it('should separate paragraphs with single newline when paragraphSpacing is false', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'First paragraph',
            runs: [{ text: 'First paragraph' }],
          },
          {
            text: 'Second paragraph',
            runs: [{ text: 'Second paragraph' }],
          },
        ],
        text: 'First paragraph\nSecond paragraph',
      };

      const result = convertDocxToTypst(doc, { paragraphSpacing: false });
      expect(result.content).toBe('First paragraph\n\nSecond paragraph');
    });

    it('should convert an empty paragraph into a blank line', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Paragraph before',
            runs: [{ text: 'Paragraph before' }],
          },
          {
            text: '',
            runs: [],
            isEmpty: true,
          },
          {
            text: 'Paragraph after',
            runs: [{ text: 'Paragraph after' }],
          },
        ],
        text: 'Paragraph before\n\nParagraph after',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Paragraph before\n\n\n\nParagraph after');
      expect(result.stats.paragraphCount).toBe(3);
    });

    it('should handle single empty paragraph document', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            isEmpty: true,
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('');
      expect(result.stats.paragraphCount).toBe(1);
    });
  });

  describe('convertDocxToTypst - document header and options', () => {
    it('should prepend header comment when includeHeader is true', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Hello world',
            runs: [{ text: 'Hello world' }],
          },
        ],
        text: 'Hello world',
      };

      const result = convertDocxToTypst(doc, { includeHeader: true });
      expect(result.content).toBe('// Generated by docx2typst\n\nHello world');
    });

    it('should not prepend header comment by default', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Hello world',
            runs: [{ text: 'Hello world' }],
          },
        ],
        text: 'Hello world',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content.startsWith('//')).toBe(false);
    });
  });

  describe('convertDocxToTypst - statistics', () => {
    it('should accurately calculate document statistics', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Doc Title',
            runs: [{ text: 'Doc Title' }],
            style: 'Heading1',
          },
          {
            text: 'Section Title',
            runs: [{ text: 'Section Title' }],
            style: 'Heading2',
          },
          {
            text: 'Paragraph with formatting',
            runs: [
              { text: 'Normal text ' },
              { text: 'bold', bold: true },
              { text: ' and ' },
              { text: 'italic', italic: true },
              { text: ' and ' },
              { text: 'underlined', underline: true },
            ],
          },
          {
            text: 'Plain paragraph',
            runs: [{ text: 'Plain paragraph' }],
          },
        ],
        text: 'Doc Title\nSection Title\nParagraph with formatting\nPlain paragraph',
      };

      const result = convertDocxToTypst(doc);
      expect(result.stats).toEqual({
        paragraphCount: 2,
        headingCount: 2,
        runsWithFormatting: 3,
      });
    });

    it('should return zero statistics for empty document', () => {
      const doc: DocxDocument = {
        paragraphs: [],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.stats).toEqual({
        paragraphCount: 0,
        headingCount: 0,
        runsWithFormatting: 0,
      });
      expect(result.content).toBe('');
    });
  });

  describe('convertDocxToTypst - fallback when runs array is missing but text exists', () => {
    it('should convert paragraph using text property if runs array is empty', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Fallback text',
            runs: [],
          } as unknown as Paragraph,
        ],
        text: 'Fallback text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Fallback text');
    });

    it('should return empty string when paragraph has empty runs and text and isEmpty is false', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            isEmpty: false,
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('');
    });
  });

  describe('convertDocxToTypst - list items', () => {
    it('should convert bullet item at level 0 to "- item text"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'item text',
            runs: [{ text: 'item text' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
        ],
        text: 'item text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('- item text');
      expect(result.stats.paragraphCount).toBe(1);
    });

    it('should convert ordered item at level 0 to "+ item text"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'item text',
            runs: [{ text: 'item text' }],
            listItem: {
              abstractNumId: 1,
              numId: 2,
              level: 0,
              listType: 'ordered',
            },
          },
        ],
        text: 'item text',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('+ item text');
      expect(result.stats.paragraphCount).toBe(1);
    });

    it('should convert nested bullet at level 1 with 2 spaces to "  - nested"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'nested',
            runs: [{ text: 'nested' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 1,
              listType: 'bullet',
            },
          },
        ],
        text: 'nested',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('  - nested');
    });

    it('should convert nested ordered at level 2 with 4 spaces to "    + deep"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'deep',
            runs: [{ text: 'deep' }],
            listItem: {
              abstractNumId: 1,
              numId: 2,
              level: 2,
              listType: 'ordered',
            },
          },
        ],
        text: 'deep',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('    + deep');
    });

    it('should join two consecutive list items with "\\n" not "\\n\\n"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Item 1',
            runs: [{ text: 'Item 1' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
          {
            text: 'Item 2',
            runs: [{ text: 'Item 2' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
        ],
        text: 'Item 1\nItem 2',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('- Item 1\n\n- Item 2');
    });

    it('should convert list item with bold run to "- *bold item*"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'bold item',
            runs: [{ text: 'bold item', bold: true }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
        ],
        text: 'bold item',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('- *bold item*');
      expect(result.stats.runsWithFormatting).toBe(1);
    });

    it('should separate list paragraph followed by normal paragraph with "\\n\\n"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'List item',
            runs: [{ text: 'List item' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
          {
            text: 'Normal paragraph',
            runs: [{ text: 'Normal paragraph' }],
          },
        ],
        text: 'List item\nNormal paragraph',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('- List item\n\nNormal paragraph');
    });

    it('should separate normal paragraph followed by list paragraph with "\\n\\n"', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Normal paragraph',
            runs: [{ text: 'Normal paragraph' }],
          },
          {
            text: 'List item',
            runs: [{ text: 'List item' }],
            listItem: {
              abstractNumId: 0,
              numId: 1,
              level: 0,
              listType: 'bullet',
            },
          },
        ],
        text: 'Normal paragraph\nList item',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Normal paragraph\n\n- List item');
    });
  });

  describe('convertTableToTypst', () => {
    it('should convert a 2x2 table to correct #table() syntax', () => {
      const table: DocxTable = {
        columnCount: 2,
        rows: [
          {
            cells: [
              { paragraphs: [{ text: 'Cell 1', runs: [{ text: 'Cell 1' }] }] },
              { paragraphs: [{ text: 'Cell 2', runs: [{ text: 'Cell 2' }] }] },
            ],
          },
          {
            cells: [
              { paragraphs: [{ text: 'Cell 3', runs: [{ text: 'Cell 3' }] }] },
              { paragraphs: [{ text: 'Cell 4', runs: [{ text: 'Cell 4' }] }] },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe(
        '#table(\n  columns: 2,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); Cell 1], [#set par(spacing: 0.5em); Cell 2],\n  [#set par(spacing: 0.5em); Cell 3], [#set par(spacing: 0.5em); Cell 4],\n)',
      );
    });

    it('should format table with header row using table.header()', () => {
      const table: DocxTable = {
        columnCount: 2,
        rows: [
          {
            isHeader: true,
            cells: [
              { paragraphs: [{ text: 'Header A', runs: [{ text: 'Header A' }] }] },
              { paragraphs: [{ text: 'Header B', runs: [{ text: 'Header B' }] }] },
            ],
          },
          {
            cells: [
              { paragraphs: [{ text: 'Cell 1', runs: [{ text: 'Cell 1' }] }] },
              { paragraphs: [{ text: 'Cell 2', runs: [{ text: 'Cell 2' }] }] },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe(
        '#table(\n  columns: 2,\n  stroke: 1pt,\n  table.header(\n    [#set par(spacing: 0.5em); Header A], [#set par(spacing: 0.5em); Header B],\n  ),\n  [#set par(spacing: 0.5em); Cell 1], [#set par(spacing: 0.5em); Cell 2],\n)',
      );
    });

    it('should format table with no header using regular cell rows', () => {
      const table: DocxTable = {
        columnCount: 2,
        rows: [
          {
            cells: [
              { paragraphs: [{ text: 'Row 1 Col 1', runs: [{ text: 'Row 1 Col 1' }] }] },
              { paragraphs: [{ text: 'Row 1 Col 2', runs: [{ text: 'Row 1 Col 2' }] }] },
            ],
          },
          {
            cells: [
              { paragraphs: [{ text: 'Row 2 Col 1', runs: [{ text: 'Row 2 Col 1' }] }] },
              { paragraphs: [{ text: 'Row 2 Col 2', runs: [{ text: 'Row 2 Col 2' }] }] },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe(
        '#table(\n  columns: 2,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); Row 1 Col 1], [#set par(spacing: 0.5em); Row 1 Col 2],\n  [#set par(spacing: 0.5em); Row 2 Col 1], [#set par(spacing: 0.5em); Row 2 Col 2],\n)',
      );
    });

    it('should convert cell with bold text using *bold* in output', () => {
      const table: DocxTable = {
        columnCount: 1,
        rows: [
          {
            cells: [
              {
                paragraphs: [
                  {
                    text: 'Bold text',
                    runs: [{ text: 'Bold text', bold: true }],
                  },
                ],
              },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe('#table(\n  columns: 1,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); *Bold text*],\n)');
    });

    it('should produce [] for an empty cell', () => {
      const table: DocxTable = {
        columnCount: 2,
        rows: [
          {
            cells: [
              {
                paragraphs: [],
              },
              {
                paragraphs: [
                  {
                    text: '',
                    runs: [],
                    isEmpty: true,
                  },
                ],
              },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe('#table(\n  columns: 2,\n  stroke: 1pt,\n  [], [#set par(spacing: 0.5em); #v(1em)],\n)');
    });

    it('should format single-row table (no header)', () => {
      const table: DocxTable = {
        columnCount: 2,
        rows: [
          {
            cells: [
              { paragraphs: [{ text: 'Item 1', runs: [{ text: 'Item 1' }] }] },
              { paragraphs: [{ text: 'Item 2', runs: [{ text: 'Item 2' }] }] },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe('#table(\n  columns: 2,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); Item 1], [#set par(spacing: 0.5em); Item 2],\n)');
    });

    it('should format cell with multiple paragraphs joined with double newline', () => {
      const table: DocxTable = {
        columnCount: 1,
        rows: [
          {
            cells: [
              {
                paragraphs: [
                  { text: 'Line 1', runs: [{ text: 'Line 1' }] },
                  { text: 'Line 2', runs: [{ text: 'Line 2' }] },
                ],
              },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe('#table(\n  columns: 1,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); Line 1\n\nLine 2],\n)');
    });

    it('should append tables at end of document in convertDocxToTypst', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Intro paragraph',
            runs: [{ text: 'Intro paragraph' }],
          },
        ],
        text: 'Intro paragraph',
        tables: [
          {
            columnCount: 2,
            rows: [
              {
                cells: [
                  { paragraphs: [{ text: 'A', runs: [{ text: 'A' }] }] },
                  { paragraphs: [{ text: 'B', runs: [{ text: 'B' }] }] },
                ],
              },
            ],
          },
        ],
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe(
        'Intro paragraph\n\n#table(\n  columns: 2,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); A], [#set par(spacing: 0.5em); B],\n)',
      );
    });

    it('should convert document containing only tables and no paragraphs', () => {
      const doc: DocxDocument = {
        paragraphs: [],
        text: '',
        tables: [
          {
            columnCount: 1,
            rows: [
              {
                cells: [
                  { paragraphs: [{ text: 'Solo cell', runs: [{ text: 'Solo cell' }] }] },
                ],
              },
            ],
          },
        ],
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#table(\n  columns: 1,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); Solo cell],\n)');
    });
  });

  describe('convertDocxToTypst - images', () => {
    it('should produce #image("images/image1.png") for paragraph with image', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/image1.png',
                zipPath: 'word/media/image1.png',
                mimeType: 'image/png',
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#image("images/image1.png")');
    });

    it('should produce #image("images/image1.png", width: 50%) for image with widthEmu', () => {
      // 5486400 * 0.5 = 2743200 EMUs
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/image1.png',
                zipPath: 'word/media/image1.png',
                widthEmu: 2743200,
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#image("images/image1.png", width: 50%)');
    });

    it('should use custom imageOutputDir option when provided', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/image1.png',
                zipPath: 'word/media/image1.png',
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc, { imageOutputDir: 'custom_images' });
      expect(result.content).toBe('#image("custom_images/image1.png")');
    });

    it('should cap width percentage at 100% when width > page width', () => {
      // 6000000 > 5486400
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/image1.png',
                zipPath: 'word/media/image1.png',
                widthEmu: 6000000,
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#image("images/image1.png", width: 100%)');
    });

    it('should append image markup on new line when paragraph has text', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Diagram caption',
            runs: [{ text: 'Diagram caption' }],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/diagram.png',
                zipPath: 'word/media/diagram.png',
              },
            ],
          },
        ],
        text: 'Diagram caption',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('Diagram caption\n#image("images/diagram.png")');
    });

    it('should emit multiple images on separate lines', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/img1.png',
                zipPath: 'word/media/img1.png',
              },
              {
                relationshipId: 'rId2',
                targetPath: 'media/img2.png',
                zipPath: 'word/media/img2.png',
                widthEmu: 1371600,
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe(
        '#image("images/img1.png")\n#image("images/img2.png", width: 25%)',
      );
    });

    it('should handle custom imageOutputDir with trailing slash', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [],
            images: [
              {
                relationshipId: 'rId1',
                targetPath: 'media/image1.png',
                zipPath: 'word/media/image1.png',
              },
            ],
          },
        ],
        text: '',
      };

      const result = convertDocxToTypst(doc, { imageOutputDir: 'my_images/' });
      expect(result.content).toBe('#image("my_images/image1.png")');
    });
  });

  describe('FIX: Text Color, Table Borders, and Page Breaks', () => {
    it('should convert colored runs to #text(fill: rgb("..."))[...]', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Warning text',
            runs: [{ text: 'Warning text', color: 'ff0000' }],
          },
        ],
        text: 'Warning text',
      };
      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#text(fill: rgb("ff0000"))[Warning text]');
    });

    it('should wrap bold, italic, and underline inside #text(fill: rgb("..."))', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'styled',
            runs: [
              {
                text: 'styled',
                bold: true,
                italic: true,
                underline: true,
                color: 'ff0000',
              },
            ],
          },
        ],
        text: 'styled',
      };
      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#text(fill: rgb("ff0000"))[#underline[*_styled_*]]');
    });

    it('should render horizontal-only tables with stroke: none and table.hline', () => {
      const table: DocxTable = {
        columnCount: 2,
        borders: {
          top: true,
          bottom: true,
          left: false,
          right: false,
          insideH: true,
          insideV: false,
        },
        rows: [
          {
            cells: [
              { paragraphs: [{ text: 'A', runs: [{ text: 'A' }] }] },
              { paragraphs: [{ text: 'B', runs: [{ text: 'B' }] }] },
            ],
          },
        ],
      };
      const result = convertTableToTypst(table);
      expect(result).toBe(
        '#table(\n  columns: 2,\n  stroke: (x: none, y: 1pt),\n  [#set par(spacing: 0.5em); A], [#set par(spacing: 0.5em); B],\n)',
      );
    });

    it('should render full-grid tables with stroke: 1pt when all borders are present', () => {
      const table: DocxTable = {
        columnCount: 2,
        borders: {
          top: true,
          bottom: true,
          left: true,
          right: true,
          insideH: true,
          insideV: true,
        },
        rows: [
          {
            cells: [
              { paragraphs: [{ text: 'A', runs: [{ text: 'A' }] }] },
              { paragraphs: [{ text: 'B', runs: [{ text: 'B' }] }] },
            ],
          },
        ],
      };
      const result = convertTableToTypst(table);
      expect(result).toBe('#table(\n  columns: 2,\n  stroke: 1pt,\n  [#set par(spacing: 0.5em); A], [#set par(spacing: 0.5em); B],\n)');
    });

    it('should handle pageBreakBefore on paragraph', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Page 2 content',
            runs: [{ text: 'Page 2 content' }],
            pageBreakBefore: true,
          },
        ],
        text: 'Page 2 content',
      };
      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#pagebreak()\nPage 2 content');
    });

    it('should handle sectionBreak pageBreak on paragraph', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Section 1 content',
            runs: [{ text: 'Section 1 content' }],
            sectionBreak: { columnCount: 1, pageBreak: true },
          },
          {
            text: 'Section 2 content',
            runs: [{ text: 'Section 2 content' }],
          },
        ],
        text: 'Section 1 content\nSection 2 content',
      };
      const result = convertDocxToTypst(doc);
      expect(result.content).toContain('#pagebreak()');
      expect(result.content).toBe(
        'Section 1 content\n#pagebreak()\n\nSection 2 content',
      );
    });

    it('should handle inline pageBreak on TextRun', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: '',
            runs: [{ text: '', pageBreak: true }],
          },
        ],
        text: '',
      };
      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#pagebreak()');
    });
  });

  describe('Paragraph Indentation', () => {
    describe('twipsToCm', () => {
      it('should correctly convert twips to centimeters formatted for Typst', () => {
        expect(twipsToCm(1440)).toBe('2.54cm');
        expect(twipsToCm(720)).toBe('1.27cm');
        expect(twipsToCm(567)).toBe('1.00cm');
        expect(twipsToCm(709)).toBe('1.25cm');
        expect(twipsToCm(284)).toBe('0.50cm');
        expect(twipsToCm(425)).toBe('0.75cm');
        expect(twipsToCm(851)).toBe('1.50cm');
        expect(twipsToCm(993)).toBe('1.75cm');
        expect(twipsToCm(1134)).toBe('2.00cm');
      });
    });

    describe('first-line indent', () => {
      it('should prepend #h(X.XXcm) for first-line indent on normal paragraph', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'The introduction contains background problems.',
              runs: [{ text: 'The introduction contains background problems.' }],
              indent: { firstLine: 720 },
            },
          ],
          text: 'The introduction contains background problems.',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('#h(1.27cm)The introduction contains background problems.');
      });
    });

    describe('hanging indent', () => {
      it('should wrap in #pad(left: Xcm)[#h(-Ycm)content] when both left and hanging are set', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Baxter, C. 1997. Race equality in health care and education.',
              runs: [{ text: 'Baxter, C. 1997. Race equality in health care and education.' }],
              indent: { left: 993, hanging: 425 },
            },
          ],
          text: 'Baxter, C. 1997. Race equality in health care and education.',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe(
          '#pad(left: 1.75cm)[#h(-0.75cm)Baxter, C. 1997. Race equality in health care and education.]',
        );
      });

      it('should use hanging as left pad when left is not explicitly specified', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Hanging without explicit left',
              runs: [{ text: 'Hanging without explicit left' }],
              indent: { hanging: 567 },
            },
          ],
          text: 'Hanging without explicit left',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe(
          '#pad(left: 1.00cm)[#h(-1.00cm)Hanging without explicit left]',
        );
      });
    });

    describe('left and right indent', () => {
      it('should wrap in #pad(left: Xcm)[content] when left indent is set without firstLine or hanging', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Figure 1. Risk Pathway',
              runs: [{ text: 'Figure 1. Risk Pathway' }],
              indent: { left: 567 },
            },
          ],
          text: 'Figure 1. Risk Pathway',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('#pad(left: 1.00cm)[Figure 1. Risk Pathway]');
      });

      it('should wrap in #pad(right: Xcm)[content] when right indent is set', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Right indented paragraph',
              runs: [{ text: 'Right indented paragraph' }],
              indent: { right: 567 },
            },
          ],
          text: 'Right indented paragraph',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('#pad(right: 1.00cm)[Right indented paragraph]');
      });

      it('should combine left and right indent in single #pad', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Both margins indented',
              runs: [{ text: 'Both margins indented' }],
              indent: { left: 567, right: 284 },
            },
          ],
          text: 'Both margins indented',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('#pad(left: 1.00cm, right: 0.50cm)[Both margins indented]');
      });

      it('should combine left indent and first-line indent', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Block quote with first-line indent',
              runs: [{ text: 'Block quote with first-line indent' }],
              indent: { left: 567, firstLine: 284 },
            },
          ],
          text: 'Block quote with first-line indent',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('#pad(left: 1.00cm)[#h(0.50cm)Block quote with first-line indent]');
      });
    });

    describe('indentation exclusions', () => {
      it('should not apply indentation to headings', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Introduction Heading',
              runs: [{ text: 'Introduction Heading' }],
              style: 'Heading1',
              indent: { firstLine: 720, left: 567 },
            },
          ],
          text: 'Introduction Heading',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('= Introduction Heading');
      });

      it('should not apply indentation to list items', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'List item with indent in docx',
              runs: [{ text: 'List item with indent in docx' }],
              listItem: {
                numId: 1,
                level: 0,
                abstractNumId: 1,
                listType: 'bullet',
              },
              indent: { left: 720, hanging: 360 },
            },
          ],
          text: 'List item with indent in docx',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('- List item with indent in docx');
      });

      it('should not apply indentation to empty paragraphs', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: '',
              runs: [],
              isEmpty: true,
              indent: { firstLine: 720 },
            },
          ],
          text: '',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe('');
      });
    });

    describe('wrapping order with alignment', () => {
      it('should wrap alignment outside pad and first-line indent', () => {
        const doc: DocxDocument = {
          paragraphs: [
            {
              text: 'Centered indented paragraph',
              runs: [{ text: 'Centered indented paragraph' }],
              align: 'center',
              indent: { left: 567, firstLine: 284 },
            },
          ],
          text: 'Centered indented paragraph',
        };
        const result = convertDocxToTypst(doc);
        expect(result.content).toBe(
          '#align(center)[#pad(left: 1.00cm)[#h(0.50cm)Centered indented paragraph]]',
        );
      });
    });

    describe('table cell indentation', () => {
      it('should apply first-line indentation to paragraphs inside table cells', () => {
        const table: DocxTable = {
          columnCount: 1,
          rows: [
            {
              cells: [
                {
                  paragraphs: [
                    {
                      text: 'Abstract body paragraph',
                      runs: [{ text: 'Abstract body paragraph' }],
                      indent: { firstLine: 601 },
                    },
                  ],
                },
              ],
            },
          ],
        };
        const result = convertTableToTypst(table);
        expect(result).toContain('#h(1.06cm)Abstract body paragraph');
      });
    });
  });
});




