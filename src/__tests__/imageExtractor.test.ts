/**
 * @file imageExtractor.test.ts
 * Unit tests for image extraction utilities and OOXML DrawingML parsing.
 */

import { mkdir, writeFile } from 'node:fs/promises';
import { parseXml } from '../utils/xmlParser.js';
import {
  findImagesInRun,
  extractImageMetadata,
  saveImages,
} from '../extractor/imageExtractor.js';
import type { DocxImage } from '../types.js';

jest.mock('node:fs/promises', () => ({
  mkdir: jest.fn().mockResolvedValue(undefined),
  writeFile: jest.fn().mockResolvedValue(undefined),
}));

describe('imageExtractor', () => {
  const ooxmlRunWithInline = `
    <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:drawing>
        <wp:inline>
          <wp:extent cx="2743200" cy="1828800"/>
          <wp:docPr id="1" name="Picture 1" descr="A scenic view"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:blipFill>
                  <a:blip r:embed="rId1"/>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  `;

  const ooxmlRunWithAnchor = `
    <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:drawing>
        <wp:anchor>
          <wp:extent cx="1000000" cy="500000"/>
          <wp:docPr id="2" title="Anchored Image Title"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:blipFill>
                  <a:blip r:embed="rId2"/>
                </pic:blipFill>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </w:r>
  `;

  describe('findImagesInRun', () => {
    it('should find relationship IDs with OOXML run XML containing w:drawing > wp:inline > a:graphic chain', () => {
      const parsed = parseXml(ooxmlRunWithInline);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const relIds = findImagesInRun(rNode);
      expect(relIds).toEqual(['rId1']);
    });

    it('should find relationship IDs with OOXML run XML containing w:drawing > wp:anchor alternative', () => {
      const parsed = parseXml(ooxmlRunWithAnchor);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const relIds = findImagesInRun(rNode);
      expect(relIds).toEqual(['rId2']);
    });

    it('should return an empty array with no drawing element', () => {
      const rNode = {
        'w:t': 'Plain text without images',
      };
      expect(findImagesInRun(rNode)).toEqual([]);
    });

    it('should handle non-record input gracefully', () => {
      expect(findImagesInRun(null as unknown as Record<string, unknown>)).toEqual([]);
      expect(findImagesInRun(undefined as unknown as Record<string, unknown>)).toEqual([]);
      expect(findImagesInRun([] as unknown as Record<string, unknown>)).toEqual([]);
    });

    it('should navigate run structures using non-namespaced keys', () => {
      const directObj: Record<string, unknown> = {
        drawing: {
          inline: {
            graphic: {
              graphicData: {
                pic: {
                  blipFill: {
                    blip: {
                      embed: 'rId3',
                    },
                  },
                },
              },
            },
          },
        },
      };
      expect(findImagesInRun(directObj)).toEqual(['rId3']);
    });
  });

  describe('extractImageMetadata', () => {
    const baseMap = new Map<string, DocxImage>([
      [
        'rId1',
        {
          relationshipId: 'rId1',
          targetPath: 'media/image1.png',
          zipPath: 'word/media/image1.png',
          mimeType: 'image/png',
        },
      ],
      [
        'rId2',
        {
          relationshipId: 'rId2',
          targetPath: 'media/image2.jpeg',
          zipPath: 'word/media/image2.jpeg',
          mimeType: 'image/jpeg',
        },
      ],
    ]);

    it('should return correct DocxImage with relationship ID from mocked imageMap', () => {
      const parsed = parseXml(ooxmlRunWithInline);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const results = extractImageMetadata(rNode, baseMap);

      expect(results).toHaveLength(1);
      expect(results[0].relationshipId).toBe('rId1');
      expect(results[0].targetPath).toBe('media/image1.png');
      expect(results[0].zipPath).toBe('word/media/image1.png');
      expect(results[0].mimeType).toBe('image/png');
    });

    it('should extract extent attributes as widthEmu and heightEmu and altText', () => {
      const parsed = parseXml(ooxmlRunWithInline);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const results = extractImageMetadata(rNode, baseMap);

      expect(results).toHaveLength(1);
      expect(results[0].widthEmu).toBe(2743200);
      expect(results[0].heightEmu).toBe(1828800);
      expect(results[0].altText).toBe('A scenic view');
    });

    it('should extract title as altText fallback when descr is missing', () => {
      const parsed = parseXml(ooxmlRunWithAnchor);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const results = extractImageMetadata(rNode, baseMap);

      expect(results).toHaveLength(1);
      expect(results[0].relationshipId).toBe('rId2');
      expect(results[0].widthEmu).toBe(1000000);
      expect(results[0].heightEmu).toBe(500000);
      expect(results[0].altText).toBe('Anchored Image Title');
    });

    it('should return empty array when rId not in imageMap', () => {
      const parsed = parseXml(ooxmlRunWithInline);
      const rNode = parsed['w:r'] as Record<string, unknown>;
      const emptyMap = new Map<string, DocxImage>();

      const results = extractImageMetadata(rNode, emptyMap);
      expect(results).toEqual([]);

      const unrelatedMap = new Map<string, DocxImage>([
        [
          'rId99',
          {
            relationshipId: 'rId99',
            targetPath: 'media/other.png',
            zipPath: 'word/media/other.png',
          },
        ],
      ]);
      expect(extractImageMetadata(rNode, unrelatedMap)).toEqual([]);
    });

    it('should return deep copy and not mutate entries in imageMap', () => {
      const parsed = parseXml(ooxmlRunWithInline);
      const rNode = parsed['w:r'] as Record<string, unknown>;

      const originalEntry = baseMap.get('rId1')!;
      expect(originalEntry.widthEmu).toBeUndefined();
      expect(originalEntry.heightEmu).toBeUndefined();
      expect(originalEntry.altText).toBeUndefined();

      const results = extractImageMetadata(rNode, baseMap);
      expect(results[0].widthEmu).toBe(2743200);

      // Verify the map entry was NOT mutated
      expect(originalEntry.widthEmu).toBeUndefined();
      expect(originalEntry.heightEmu).toBeUndefined();
      expect(originalEntry.altText).toBeUndefined();
    });

    it('should return empty array for invalid rNode or missing imageMap', () => {
      expect(
        extractImageMetadata(
          null as unknown as Record<string, unknown>,
          baseMap,
        ),
      ).toEqual([]);
      expect(
        extractImageMetadata(
          { 'w:t': 'no drawing' },
          baseMap,
        ),
      ).toEqual([]);
    });
  });

  describe('saveImages', () => {
    beforeEach(() => {
      jest.clearAllMocks();
    });

    it('should extract binary data and save images to output directory', async () => {
      const mockZip = {
        file: jest.fn().mockImplementation((path: string) => {
          if (path === 'word/media/image1.png') {
            return {
              async: jest.fn().mockResolvedValue(new Uint8Array([1, 2, 3, 4])),
            };
          }
          return null;
        }),
      };

      const images: DocxImage[] = [
        {
          relationshipId: 'rId1',
          targetPath: 'media/image1.png',
          zipPath: 'word/media/image1.png',
          mimeType: 'image/png',
        },
      ];

      const result = await saveImages(images, mockZip, 'custom/images');

      expect(mkdir).toHaveBeenCalledWith('custom/images', { recursive: true });
      expect(writeFile).toHaveBeenCalledWith(
        'custom/images/image1.png',
        new Uint8Array([1, 2, 3, 4]),
      );
      expect(result.get('word/media/image1.png')).toBe('custom/images/image1.png');
    });

    it('should return empty map when images array is empty', async () => {
      const mockZip = { file: jest.fn() };
      const result = await saveImages([], mockZip, 'custom/images');
      expect(result.size).toBe(0);
      expect(mkdir).not.toHaveBeenCalled();
      expect(writeFile).not.toHaveBeenCalled();
    });

    it('should return empty map when zipInstance is invalid', async () => {
      const images: DocxImage[] = [
        {
          relationshipId: 'rId1',
          targetPath: 'media/image1.png',
          zipPath: 'word/media/image1.png',
        },
      ];
      const result = await saveImages(images, null, 'custom/images');
      expect(result.size).toBe(0);
    });

    it('should skip images not found in the zip archive', async () => {
      const mockZip = {
        file: jest.fn().mockReturnValue(null),
      };

      const images: DocxImage[] = [
        {
          relationshipId: 'rId1',
          targetPath: 'media/missing.png',
          zipPath: 'word/media/missing.png',
        },
      ];

      const result = await saveImages(images, mockZip, 'custom/images');
      expect(result.size).toBe(0);
      expect(writeFile).not.toHaveBeenCalled();
    });
  });
});
