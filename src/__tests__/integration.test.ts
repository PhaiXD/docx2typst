/**
 * @file integration.test.ts
 * End-to-end integration tests creating and converting real in-memory DOCX archives.
 */

import { randomUUID } from 'node:crypto';
import { unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { convertDocxFile, convertDocxToText } from '../index.js';

describe('End-to-End DOCX Conversion Integration Tests', () => {
  const createdTempFiles: string[] = [];

  /**
   * Helper to pack a real DOCX ZIP archive in memory and write it to a temporary file.
   *
   * @param documentXml - The XML content for word/document.xml.
   * @param numberingXml - Optional XML content for word/numbering.xml.
   * @returns Path to the written temporary .docx file.
   */
  async function createTempDocx(
    documentXml: string,
    numberingXml?: string,
  ): Promise<string> {
    const zip = new JSZip();
    zip.file('word/document.xml', documentXml);

    if (numberingXml) {
      zip.file('word/numbering.xml', numberingXml);
    }

    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const tempPath = path.join(tmpdir(), `docx2typst-test-${randomUUID()}.docx`);
    await writeFile(tempPath, buffer);
    createdTempFiles.push(tempPath);
    return tempPath;
  }

  afterEach(async () => {
    while (createdTempFiles.length > 0) {
      const file = createdTempFiles.pop();
      if (file) {
        try {
          await unlink(file);
        } catch {
          // Ignore cleanup errors
        }
      }
    }
  });

  afterAll(async () => {
    for (const file of createdTempFiles) {
      try {
        await unlink(file);
      } catch {
        // Ignore cleanup errors
      }
    }
  });

  it('should convert a minimal document.xml with one paragraph to Typst markup', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const tempFile = await createTempDocx(documentXml);
    const result = await convertDocxFile(tempFile);

    expect(result.content).toContain('Hello World');
    expect(result.stats.paragraphCount).toBe(1);
    expect(result.stats.headingCount).toBe(0);

    const plainText = await convertDocxToText(tempFile);
    expect(plainText).toBe('Hello World');
  });

  it('should convert a document with a Heading1 style to "= Heading"', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Heading Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Body paragraph content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const tempFile = await createTempDocx(documentXml);
    const result = await convertDocxFile(tempFile);

    expect(result.content).toContain('= Heading Title');
    expect(result.content).toContain('Body paragraph content');
    expect(result.stats.headingCount).toBe(1);
    expect(result.stats.paragraphCount).toBe(1);
  });

  it('should convert a document with a bullet list to "- "', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>First bullet item</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>Second bullet item</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`;

    const tempFile = await createTempDocx(documentXml, numberingXml);
    const result = await convertDocxFile(tempFile);

    expect(result.content).toContain('- First bullet item');
    expect(result.content).toContain('- Second bullet item');
  });

  it('should convert the fixture file from test/fixtures/minimal.docx', async () => {
    const fixturePath = path.resolve(process.cwd(), 'test', 'fixtures', 'minimal.docx');
    const result = await convertDocxFile(fixturePath);

    expect(result.content).toContain('Hello from docx2typst fixture');
    expect(result.stats.paragraphCount).toBe(1);
  });
});
