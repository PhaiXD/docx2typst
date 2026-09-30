/**
 * @file reader.test.ts
 * Unit tests for reading .docx packages and error handling.
 */

import { writeFile, unlink, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import JSZip from 'jszip';
import { readDocxFile } from '../reader.js';
import { DocxReadError } from '../errors.js';

describe('reader', () => {
  const testDir = join(tmpdir(), 'docx2typst-test-' + Date.now());
  const validDocxPath = join(testDir, 'sample.docx');
  const numberingDocxPath = join(testDir, 'with-numbering.docx');
  const invalidZipPath = join(testDir, 'corrupt.docx');
  const missingXmlDocxPath = join(testDir, 'missing-doc.docx');

  beforeAll(async () => {
    await mkdir(testDir, { recursive: true });

    // 1. Create a valid docx with word/document.xml and word/_rels/document.xml.rels
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>');
    zip.file('word/_rels/document.xml.rels', '<Relationships></Relationships>');
    const validBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    await writeFile(validDocxPath, validBuffer);

    // 1b. Create a valid docx with word/document.xml, word/_rels, and word/numbering.xml
    const numberingZip = new JSZip();
    numberingZip.file('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>');
    numberingZip.file('word/numbering.xml', '<w:numbering><w:abstractNum w:abstractNumId="0"/></w:numbering>');
    const numberingBuffer = await numberingZip.generateAsync({ type: 'nodebuffer' });
    await writeFile(numberingDocxPath, numberingBuffer);

    // 2. Create an invalid non-zip file
    await writeFile(invalidZipPath, 'NOT A ZIP FILE');

    // 3. Create a zip without word/document.xml
    const emptyZip = new JSZip();
    emptyZip.file('other.xml', '<foo/>');
    const emptyBuffer = await emptyZip.generateAsync({ type: 'nodebuffer' });
    await writeFile(missingXmlDocxPath, emptyBuffer);
  });

  afterAll(async () => {
    try {
      await unlink(validDocxPath);
      await unlink(numberingDocxPath);
      await unlink(invalidZipPath);
      await unlink(missingXmlDocxPath);
    } catch {
      // Ignore cleanup errors
    }
  });

  it('should successfully read a valid docx file', async () => {
    const result = await readDocxFile(validDocxPath);
    expect(result.documentXml).toContain('Hello');
    expect(result.relationshipsXml).toContain('Relationships');
    expect(result.numberingXml).toBeUndefined();
  });

  it('should successfully read numberingXml when present in docx', async () => {
    const result = await readDocxFile(numberingDocxPath);
    expect(result.documentXml).toContain('Hello');
    expect(result.numberingXml).toContain('abstractNum');
  });

  it('should throw DocxReadError if file does not exist', async () => {
    await expect(readDocxFile(join(testDir, 'does-not-exist.docx'))).rejects.toThrow(
      DocxReadError,
    );
  });

  it('should throw DocxReadError if filePath is empty', async () => {
    await expect(readDocxFile('')).rejects.toThrow(DocxReadError);
  });

  it('should throw DocxReadError if file is not a valid zip archive', async () => {
    await expect(readDocxFile(invalidZipPath)).rejects.toThrow(DocxReadError);
  });

  it('should throw DocxReadError if word/document.xml is missing', async () => {
    await expect(readDocxFile(missingXmlDocxPath)).rejects.toThrow(DocxReadError);
  });
});
