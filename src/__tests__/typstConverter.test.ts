/**
 * @file typstConverter.test.ts
 * Unit tests for OOXML to Typst markup conversion layer.
 */

import { convertDocxToTypst, getHeadingLevel } from '../converter/typstConverter.js';
import type { DocxDocument, Paragraph } from '../types.js';

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
      expect(result.content).toBe('= \n\n== \n\n====== ');
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
      expect(result.content).toBe('First paragraph\nSecond paragraph');
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
      expect(result.content).toBe('- Item 1\n- Item 2');
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
});

