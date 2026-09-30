/**
 * @file bugfixes.test.ts
 * Comprehensive test suite verifying the 5 critical bug fixes:
 * 1. Images "File Not Found" relative paths in CLI and Typst compiler.
 * 2. Multi-column section layouts (w:sectPr > w:cols) and Typst #columns(N) wrapping.
 * 3. Robust nested list level, type detection, and numFmt resolution.
 * 4. Table stroke: 1pt, columnSpan (colspan), alignment, and vertical merge.
 * 5. Document-order body items (paragraphs, tables, sdt) and drawing text boxes (w:txbxContent).
 */

import { randomUUID } from 'node:crypto';
import { mkdir, readFile, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {
  convertDocxToTypst,
  convertTableToTypst,
  extractNumberingMaps,
  extractText,
  parseXmlPreserveOrder,
  resolveListItems,
} from '../index.js';
import { runCli } from '../cli.js';
import type { DocxDocument, DocxTable, Paragraph } from '../types.js';

describe('Critical Bug Fixes Verification', () => {
  const tempFiles: string[] = [];

  afterEach(async () => {
    while (tempFiles.length > 0) {
      const file = tempFiles.pop();
      if (file) {
        try {
          await unlink(file);
        } catch {
          // ignore
        }
      }
    }
  });

  // ==============================================================
  // BUG 1: Images "File Not Found" in Typst Compiler
  // ==============================================================
  describe('BUG 1: Image extraction & relative path in CLI', () => {
    it('should save images to the directory relative to the output .typ file and emit relative #image path', async () => {
      const subDir = path.join(tmpdir(), `bug1-test-${randomUUID()}`);
      await mkdir(subDir, { recursive: true });

      const zip = new JSZip();
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="1828800" cy="1828800"/>
            <wp:docPr id="1" name="Picture 1" descr="Sample Image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill>
                    <a:blip r:embed="rIdImg1"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

      const fakePng = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

      zip.file('word/document.xml', docXml);
      zip.file('word/_rels/document.xml.rels', relsXml);
      zip.file('word/media/image1.png', fakePng);

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const docxPath = path.join(subDir, 'nested-doc.docx');
      const expectedTypPath = path.join(subDir, 'nested-doc.typ');
      const expectedImgDir = path.join(subDir, 'custom-images');
      const expectedImgFile = path.join(expectedImgDir, 'image1.png');

      await writeFile(docxPath, buffer);
      tempFiles.push(docxPath, expectedTypPath, expectedImgFile);

      const exitCode = await runCli(['--images-dir', 'custom-images', docxPath]);
      expect(exitCode).toBe(0);

      // Verify .typ content has relative path
      const typContent = await readFile(expectedTypPath, 'utf-8');
      expect(typContent).toContain('#image("custom-images/image1.png"');

      // Verify image was saved to the absolute directory next to the .typ file
      const savedImgData = await readFile(expectedImgFile);
      expect(savedImgData).toEqual(fakePng);
    });
  });

  // ==============================================================
  // BUG 2: Multi-Column Layout (w:sectPr > w:cols)
  // ==============================================================
  describe('BUG 2: Multi-column section layouts', () => {
    it('should extract sectionBreak when paragraph pPr contains w:sectPr with cols > 1', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:cols w:num="2" w:space="720"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>Section 1 Paragraph</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Section 2 Paragraph</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(2);
      expect(doc.paragraphs[0].sectionBreak).toBeDefined();
      expect(doc.paragraphs[0].sectionBreak?.columnCount).toBe(2);
      expect(doc.paragraphs[1].sectionBreak).toBeUndefined();
    });

    it('should extract final section properties from body w:sectPr', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Only Paragraph</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:cols w:num="3" w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.sections).toBeDefined();
      expect(doc.sections?.[0].columnCount).toBe(3);
    });

    it('should wrap 2-column sections in #columns(2) with 2-space indentation', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'First column content',
            runs: [{ text: 'First column content' }],
          },
          {
            text: 'Second column content',
            runs: [{ text: 'Second column content' }],
            sectionBreak: { columnCount: 2 },
          },
          {
            text: 'Single column body paragraph',
            runs: [{ text: 'Single column body paragraph' }],
          },
        ],
        text: 'First column content\nSecond column content\nSingle column body paragraph',
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe(
        '#columns(2)[\n  First column content\n\n  Second column content\n]\n\nSingle column body paragraph',
      );
    });

    it('should wrap document in #columns when final section has columnCount > 1', () => {
      const doc: DocxDocument = {
        paragraphs: [
          {
            text: 'Intro in 2 columns',
            runs: [{ text: 'Intro in 2 columns' }],
          },
        ],
        text: 'Intro in 2 columns',
        sections: [{ columnCount: 2 }],
      };

      const result = convertDocxToTypst(doc);
      expect(result.content).toBe('#columns(2)[\n  Intro in 2 columns\n]');
    });
  });

  // ==============================================================
  // BUG 3: Broken Nested Lists - Level and Type Detection
  // ==============================================================
  describe('BUG 3: Nested lists and numbering formats', () => {
    it('should extract abstractNumId case-insensitively and parse abstractNumFmtMap', () => {
      const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="100">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="lowerLetter"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="lowerRoman"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="100"/>
  </w:num>
</w:numbering>`;

      const { numIdMap, abstractNumMap, abstractNumFmtMap } = extractNumberingMaps(numberingXml);
      expect(numIdMap.get(1)).toBe(100);
      expect(abstractNumMap.get('100:0')).toBe('ordered');
      expect(abstractNumMap.get('100:1')).toBe('ordered');
      expect(abstractNumMap.get('100:2')).toBe('ordered');

      expect(abstractNumFmtMap.get('100:0')).toBe('decimal');
      expect(abstractNumFmtMap.get('100:1')).toBe('lowerletter');
      expect(abstractNumFmtMap.get('100:2')).toBe('lowerroman');
    });

    it('should handle abstractNum node missing id gracefully and warn', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
      const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum>
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
    </w:lvl>
  </w:abstractNum>
</w:numbering>`;

      const { numIdMap, abstractNumMap } = extractNumberingMaps(numberingXml);
      expect(numIdMap.size).toBe(0);
      expect(abstractNumMap.size).toBe(0);
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('Warning: Could not extract abstractNumId'),
        expect.anything(),
      );
      warnSpy.mockRestore();
    });

    it('should resolve numFmt in resolveListItems and keep bodyItems in sync', () => {
      const p1: Paragraph = {
        text: 'Decimal Item',
        runs: [{ text: 'Decimal Item' }],
        listItem: { numId: 1, level: 0, abstractNumId: 0, listType: 'bullet' },
      };
      const p2: Paragraph = {
        text: 'Letter Item',
        runs: [{ text: 'Letter Item' }],
        listItem: { numId: 1, level: 1, abstractNumId: 0, listType: 'bullet' },
      };

      const doc: DocxDocument = {
        paragraphs: [p1, p2],
        text: 'Decimal Item\nLetter Item',
        bodyItems: [
          { type: 'paragraph', paragraph: p1 },
          { type: 'paragraph', paragraph: p2 },
        ],
      };

      const numIdMap = new Map([[1, 5]]);
      const abstractNumMap = new Map([
        ['5:0', 'ordered' as const],
        ['5:1', 'ordered' as const],
      ]);
      const abstractNumFmtMap = new Map([
        ['5:0', 'decimal'],
        ['5:1', 'lowerLetter'],
      ]);

      const resolved = resolveListItems(doc, numIdMap, abstractNumMap, abstractNumFmtMap);

      expect(resolved.paragraphs[0].listItem?.listType).toBe('ordered');
      expect(resolved.paragraphs[0].listItem?.numFmt).toBe('decimal');
      expect(resolved.paragraphs[1].listItem?.listType).toBe('ordered');
      expect(resolved.paragraphs[1].listItem?.numFmt).toBe('lowerLetter');

      // Check bodyItems were also updated
      expect(resolved.bodyItems?.[0]).toEqual({
        type: 'paragraph',
        paragraph: resolved.paragraphs[0],
      });
      expect(resolved.bodyItems?.[1]).toEqual({
        type: 'paragraph',
        paragraph: resolved.paragraphs[1],
      });

      // Verify typst output produces + with 2-space indentation per level
      const typstDoc = convertDocxToTypst(resolved);
      expect(typstDoc.content).toBe('+ Decimal Item\n  + Letter Item');
    });
  });

  // ==============================================================
  // BUG 4: Tables Missing Borders and Incorrect Column Span
  // ==============================================================
  describe('BUG 4: Tables stroke, colspan, alignment, and vertical merge', () => {
    it('should extract cell alignment from w:tcPr > w:jc or paragraph w:jc', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:jc w:val="center"/>
          </w:tcPr>
          <w:p><w:r><w:t>Centered Cell</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:pPr>
              <w:jc w:val="right"/>
            </w:pPr>
            <w:r><w:t>Right Cell</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.tables).toHaveLength(1);
      const cells = doc.tables![0].rows[0].cells;
      expect(cells[0].align).toBe('center');
      expect(cells[1].align).toBe('right');
    });

    it('should convert table with stroke: 1pt, colspan, align, and vertical merge', () => {
      const table: DocxTable = {
        columnCount: 3,
        rows: [
          {
            cells: [
              {
                columnSpan: 2,
                align: 'center',
                paragraphs: [{ text: 'Spanned 2 Center', runs: [{ text: 'Spanned 2 Center' }] }],
              },
              {
                align: 'right',
                paragraphs: [{ text: 'Right', runs: [{ text: 'Right' }] }],
              },
            ],
          },
          {
            cells: [
              {
                isVerticalMerge: true,
                paragraphs: [{ text: 'Continuation', runs: [{ text: 'Continuation' }] }],
              },
              {
                paragraphs: [{ text: 'Normal', runs: [{ text: 'Normal' }] }],
              },
              {
                paragraphs: [],
              },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toBe(
        '#table(\n' +
          '  columns: 3,\n' +
          '  stroke: 1pt,\n' +
          '  table.cell(colspan: 2, align: center)[Spanned 2 Center], table.cell(align: right)[Right],\n' +
          '  [], [Normal], [],\n' +
          ')',
      );
    });
  });

  // ==============================================================
  // BUG 5: Floating Elements and Text Boxes Out of Order
  // ==============================================================
  describe('BUG 5: Document order and text boxes', () => {
    it('should parse preserveOrder XML array using parseXmlPreserveOrder', () => {
      const xml = '<w:document><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>';
      const ordered = parseXmlPreserveOrder(xml);
      expect(Array.isArray(ordered)).toBe(true);
      expect(ordered.length).toBeGreaterThan(0);
    });

    it('should preserve interleaved order of paragraphs and tables via bodyItems', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Paragraph 1</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Table Cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>Paragraph 2</w:t></w:r></w:p>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.bodyItems).toHaveLength(3);
      expect(doc.bodyItems?.[0].type).toBe('paragraph');
      expect((doc.bodyItems?.[0] as { paragraph: Paragraph }).paragraph.text).toBe('Paragraph 1');
      expect(doc.bodyItems?.[1].type).toBe('table');
      expect(doc.bodyItems?.[2].type).toBe('paragraph');
      expect((doc.bodyItems?.[2] as { paragraph: Paragraph }).paragraph.text).toBe('Paragraph 2');

      const typst = convertDocxToTypst(doc);
      expect(typst.content).toBe(
        'Paragraph 1\n\n#table(\n  columns: 1,\n  stroke: 1pt,\n  [Table Cell],\n)\n\nParagraph 2',
      );
    });

    it('should extract text from w:sdt content controls in document order', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before SDT</w:t></w:r></w:p>
    <w:sdt>
      <w:sdtContent>
        <w:p><w:r><w:t>Inside SDT</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p><w:r><w:t>After SDT</w:t></w:r></w:p>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(3);
      expect(doc.paragraphs[0].text).toBe('Before SDT');
      expect(doc.paragraphs[1].text).toBe('Inside SDT');
      expect(doc.paragraphs[2].text).toBe('After SDT');
      expect(doc.text).toBe('Before SDT\nInside SDT\nAfter SDT');
    });

    it('should extract text from drawing text boxes (wps:txbx > w:txbxContent)', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Leading text: </w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:anchor>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:t>Text inside floating box</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

      const doc = extractText(xml);
      expect(doc.paragraphs).toHaveLength(1);
      expect(doc.paragraphs[0].text).toContain('Leading text:');
      expect(doc.paragraphs[0].text).toContain('Text inside floating box');
    });

    it('should parse nested w:ilvl child elements in numberingExtractor', () => {
      const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="42">
    <w:lvl>
      <w:ilvl w:val="0"/>
      <w:numFmt w:val="decimal"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="7">
    <w:abstractNumId w:val="42"/>
  </w:num>
</w:numbering>`;

      const { numIdMap, abstractNumMap, abstractNumFmtMap } = extractNumberingMaps(numberingXml);
      expect(numIdMap.get(7)).toBe(42);
      expect(abstractNumMap.get('42:0')).toBe('ordered');
      expect(abstractNumFmtMap.get('42:0')).toBe('decimal');
    });

    it('should fallback to level 0 list type when sub-level is not explicitly defined in abstractNum', () => {
      const p: Paragraph = {
        text: 'Level 2 item inheriting level 0 ordered',
        runs: [{ text: 'Level 2 item inheriting level 0 ordered' }],
        listItem: { numId: 1, level: 2, abstractNumId: 0, listType: 'bullet' },
      };

      const doc: DocxDocument = {
        paragraphs: [p],
        text: p.text,
      };

      const numIdMap = new Map([[1, 99]]);
      // Only level 0 is defined in abstractNumMap
      const abstractNumMap = new Map([['99:0', 'ordered' as const]]);
      const abstractNumFmtMap = new Map([['99:0', 'decimal']]);

      const resolved = resolveListItems(doc, numIdMap, abstractNumMap, abstractNumFmtMap);
      expect(resolved.paragraphs[0].listItem?.listType).toBe('ordered');
      expect(resolved.paragraphs[0].listItem?.numFmt).toBe('decimal');
    });

    it('should render images within table cells in Typst table conversion', () => {
      const table: DocxTable = {
        columnCount: 1,
        rows: [
          {
            cells: [
              {
                paragraphs: [
                  {
                    text: 'Image caption',
                    runs: [{ text: 'Image caption' }],
                    images: [
                      {
                        relationshipId: 'rId1',
                        targetPath: 'media/cell-img.png',
                        zipPath: 'word/media/cell-img.png',
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      };

      const result = convertTableToTypst(table);
      expect(result).toContain('#image("images/cell-img.png")');
      expect(result).toContain('Image caption');
    });
  });
});
