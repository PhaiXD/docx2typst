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

  const imagesDocxPath = join(testDir, 'with-images.docx');

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

    // 1c. Create a docx with image relationships and media files
    const imagesZip = new JSZip();
    imagesZip.file('word/document.xml', '<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>');
    imagesZip.file(
      'word/_rels/document.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo.jpeg"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/anim.gif"/>
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/vector.svg"/>
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/diagram.emf"/>
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/chart.wmf"/>
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/modern.webp"/>
        <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/custom.bin"/>
        <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/missing.png"/>
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>`,
    );
    imagesZip.file('word/media/image1.png', 'fake png data');
    imagesZip.file('word/media/photo.jpeg', 'fake jpeg data');
    imagesZip.file('word/media/anim.gif', 'fake gif data');
    imagesZip.file('word/media/vector.svg', '<svg></svg>');
    imagesZip.file('word/media/diagram.emf', 'fake emf data');
    imagesZip.file('word/media/chart.wmf', 'fake wmf data');
    imagesZip.file('word/media/modern.webp', 'fake webp data');
    imagesZip.file('word/media/custom.bin', 'fake bin data');
    const imagesBuffer = await imagesZip.generateAsync({ type: 'nodebuffer' });
    await writeFile(imagesDocxPath, imagesBuffer);

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
      await unlink(imagesDocxPath);
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
    expect(result.imageMap).toBeDefined();
    expect(result.zipInstance).toBeDefined();
  });

  it('should successfully read numberingXml when present in docx', async () => {
    const result = await readDocxFile(numberingDocxPath);
    expect(result.documentXml).toContain('Hello');
    expect(result.numberingXml).toContain('abstractNum');
  });

  it('should extract images into imageMap from relationships and zip', async () => {
    const result = await readDocxFile(imagesDocxPath);
    expect(result.imageMap).toBeDefined();
    const map = result.imageMap!;

    // rId1 (.png)
    expect(map.has('rId1')).toBe(true);
    expect(map.get('rId1')).toEqual({
      relationshipId: 'rId1',
      targetPath: 'media/image1.png',
      zipPath: 'word/media/image1.png',
      mimeType: 'image/png',
    });

    // rId2 (.jpeg)
    expect(map.get('rId2')?.mimeType).toBe('image/jpeg');

    // rId3 (.gif)
    expect(map.get('rId3')?.mimeType).toBe('image/gif');

    // rId4 (.svg)
    expect(map.get('rId4')?.mimeType).toBe('image/svg+xml');

    // rId5 (.emf)
    expect(map.get('rId5')?.mimeType).toBe('image/emf');

    // rId6 (.wmf)
    expect(map.get('rId6')?.mimeType).toBe('image/wmf');

    // rId7 (.webp)
    expect(map.get('rId7')?.mimeType).toBe('image/webp');

    // rId8 (unknown extension -> default image/png)
    expect(map.get('rId8')?.mimeType).toBe('image/png');

    // rId9 (missing from zip -> should not be in imageMap)
    expect(map.has('rId9')).toBe(false);

    // rId10 (styles relationship, not image -> should not be in imageMap)
    expect(map.has('rId10')).toBe(false);
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
