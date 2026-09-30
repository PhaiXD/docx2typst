/**
 * @file textExtractor.test.ts
 * Unit tests for text extraction from OOXML markup.
 */

import { extractText, resolveListItems } from '../extractor/textExtractor.js';
import { DocxParseError } from '../errors.js';
import type { AbstractNumMap, DocxDocument, NumIdMap } from '../types.js';

describe('textExtractor', () => {
  describe('basic paragraph extraction', () => {
    it('should extract plain text from a simple paragraph', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Hello world</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(1);
      expect(doc.paragraphs[0].text).toBe('Hello world');
      expect(doc.paragraphs[0].runs).toHaveLength(1);
      expect(doc.paragraphs[0].runs[0].text).toBe('Hello world');
      expect(doc.paragraphs[0].runs[0].bold).toBeUndefined();
      expect(doc.paragraphs[0].runs[0].italic).toBeUndefined();
      expect(doc.text).toBe('Hello world');
    });

    it('should extract multiple paragraphs joined by newlines', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>First line</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r>
              <w:t>Second line</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(2);
      expect(doc.paragraphs[0].text).toBe('First line');
      expect(doc.paragraphs[1].text).toBe('Second line');
      expect(doc.text).toBe('First line\nSecond line');
    });
  });

  describe('multiple runs per paragraph', () => {
    it('should combine multiple runs in order within a single paragraph', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Hello</w:t>
            </w:r>
            <w:r>
              <w:t>brave</w:t>
            </w:r>
            <w:r>
              <w:t>world</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(1);
      expect(doc.paragraphs[0].runs).toHaveLength(3);
      expect(doc.paragraphs[0].runs[0].text).toBe('Hello');
      expect(doc.paragraphs[0].runs[1].text).toBe('brave');
      expect(doc.paragraphs[0].runs[2].text).toBe('world');
      expect(doc.paragraphs[0].text).toBe('Hellobraveworld');
    });
  });

  describe('formatting detection (bold, italic, underline)', () => {
    it('should detect bold formatting on a run', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b/>
              </w:rPr>
              <w:t>Bold text</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      const run = doc.paragraphs[0].runs[0];
      expect(run.text).toBe('Bold text');
      expect(run.bold).toBe(true);
      expect(run.italic).toBeUndefined();
      expect(run.underline).toBeUndefined();
    });

    it('should detect italic formatting on a run', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                <w:i w:val="true"/>
              </w:rPr>
              <w:t>Italic text</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      const run = doc.paragraphs[0].runs[0];
      expect(run.text).toBe('Italic text');
      expect(run.italic).toBe(true);
      expect(run.bold).toBeUndefined();
    });

    it('should detect underline formatting on a run', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:t>Underlined text</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      const run = doc.paragraphs[0].runs[0];
      expect(run.text).toBe('Underlined text');
      expect(run.underline).toBe(true);
    });

    it('should handle false/none attributes for formatting correctly', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b w:val="false"/>
                <w:i w:val="0"/>
                <w:u w:val="none"/>
              </w:rPr>
              <w:t>Plain text</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      const run = doc.paragraphs[0].runs[0];
      expect(run.bold).toBeUndefined();
      expect(run.italic).toBeUndefined();
      expect(run.underline).toBeUndefined();
    });

    it('should detect combined bold, italic, and underline on a single run', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b/>
                <w:i/>
                <w:u w:val="single"/>
              </w:rPr>
              <w:t>Formatted all</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      const run = doc.paragraphs[0].runs[0];
      expect(run.bold).toBe(true);
      expect(run.italic).toBe(true);
      expect(run.underline).toBe(true);
    });
  });

  describe('paragraph styles', () => {
    it('should extract paragraph style from w:pPr > w:pStyle @w:val', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
              <w:t>Chapter 1</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].style).toBe('Heading1');
      expect(doc.paragraphs[0].text).toBe('Chapter 1');
    });

    it('should respect includeStyles=false option', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
              <w:t>Chapter 1</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml, { includeStyles: false });
      expect(doc.paragraphs[0].style).toBeUndefined();
      expect(doc.paragraphs[0].text).toBe('Chapter 1');
    });
  });

  describe('edge cases', () => {
    it('should handle completely empty paragraphs', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p/>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(1);
      expect(doc.paragraphs[0].text).toBe('');
      expect(doc.paragraphs[0].runs).toHaveLength(0);
      expect(doc.paragraphs[0].isEmpty).toBe(true);
      expect(doc.text).toBe('');
    });

    it('should handle paragraphs with only w:pPr and no runs', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
            </w:pPr>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].text).toBe('');
      expect(doc.paragraphs[0].runs).toHaveLength(0);
      expect(doc.paragraphs[0].style).toBe('Normal');
      expect(doc.paragraphs[0].isEmpty).toBe(true);
    });

    it('should handle runs missing w:t (e.g., drawing or empty run)', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:drawing/>
            </w:r>
            <w:r>
              <w:t>Visible text</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].text).toBe('Visible text');
      expect(doc.paragraphs[0].runs).toHaveLength(2);
      expect(doc.paragraphs[0].runs[0].text).toBe('');
      expect(doc.paragraphs[0].runs[1].text).toBe('Visible text');
    });

    it('should handle xml:space="preserve" text nodes', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t xml:space="preserve">spaced content</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].text).toBe('spaced content');
    });

    it('should handle special inline elements: w:tab and w:br', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Col1</w:t>
              <w:tab/>
              <w:t>Col2</w:t>
              <w:br/>
              <w:t>Col2 next line</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].text).toContain('\t');
      expect(doc.paragraphs[0].text).toContain('\n');
    });

    it('should handle runs inside w:hyperlink elements', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:hyperlink r:id="rId4" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:r>
                <w:rPr>
                  <w:u w:val="single"/>
                </w:rPr>
                <w:t>Link text</w:t>
              </w:r>
            </w:hyperlink>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].text).toBe('Link text');
      expect(doc.paragraphs[0].runs).toHaveLength(1);
      expect(doc.paragraphs[0].runs[0].text).toBe('Link text');
      expect(doc.paragraphs[0].runs[0].underline).toBe(true);
    });

    it('should throw DocxParseError when XML is completely invalid', () => {
      expect(() => extractText('<<<not valid xml>>>')).toThrow(DocxParseError);
    });
  });

  describe('list items extraction and resolution', () => {
    it('should detect list item with w:numPr and set placeholder listItem', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:numPr>
                <w:ilvl w:val="1"/>
                <w:numId w:val="5"/>
              </w:numPr>
            </w:pPr>
            <w:r><w:t>Nested list item</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(1);
      expect(doc.paragraphs[0].listItem).toBeDefined();
      expect(doc.paragraphs[0].listItem).toEqual({
        numId: 5,
        level: 1,
        abstractNumId: 0,
        listType: 'bullet',
      });
    });

    it('should not set listItem if w:numPr has numId="0"', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="0"/>
              </w:numPr>
            </w:pPr>
            <w:r><w:t>Non-list item with numId 0</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].listItem).toBeUndefined();
    });

    it('should default level to 0 when w:ilvl is not specified', () => {
      const xml = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:numPr>
                <w:numId w:val="3"/>
              </w:numPr>
            </w:pPr>
            <w:r><w:t>Default level item</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs[0].listItem?.level).toBe(0);
      expect(doc.paragraphs[0].listItem?.numId).toBe(3);
    });

    it('should resolve list items without mutating the original document', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Bullet item',
            runs: [{ text: 'Bullet item' }],
            listItem: { numId: 1, level: 0, abstractNumId: 0, listType: 'bullet' },
          },
          {
            text: 'Numbered item',
            runs: [{ text: 'Numbered item' }],
            listItem: { numId: 2, level: 0, abstractNumId: 0, listType: 'bullet' },
          },
          {
            text: 'Regular paragraph',
            runs: [{ text: 'Regular paragraph' }],
          },
        ],
        text: 'Bullet item\nNumbered item\nRegular paragraph',
      };

      const numIdMap: NumIdMap = new Map([
        [1, 10],
        [2, 20],
      ]);
      const abstractNumMap: AbstractNumMap = new Map([
        ['10:0', 'bullet'],
        ['20:0', 'ordered'],
      ]);

      const resolved = resolveListItems(doc, numIdMap, abstractNumMap);

      // Verify original is untouched
      expect(doc.paragraphs[0].listItem?.abstractNumId).toBe(0);
      expect(doc.paragraphs[1].listItem?.listType).toBe('bullet');

      // Verify resolved document
      expect(resolved.paragraphs[0].listItem).toEqual({
        numId: 1,
        level: 0,
        abstractNumId: 10,
        listType: 'bullet',
      });
      expect(resolved.paragraphs[1].listItem).toEqual({
        numId: 2,
        level: 0,
        abstractNumId: 20,
        listType: 'ordered',
      });
      expect(resolved.paragraphs[2].listItem).toBeUndefined();
      expect(resolved.text).toBe(doc.text);
    });

    it('should default to "bullet" if abstractNumMap does not contain matching key', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Unknown list item',
            runs: [{ text: 'Unknown list item' }],
            listItem: { numId: 99, level: 3, abstractNumId: 0, listType: 'bullet' },
          },
        ],
        text: 'Unknown list item',
      };

      const resolved = resolveListItems(doc, new Map(), new Map());
      expect(resolved.paragraphs[0].listItem?.abstractNumId).toBe(0);
      expect(resolved.paragraphs[0].listItem?.listType).toBe('bullet');
    });
  });
});
