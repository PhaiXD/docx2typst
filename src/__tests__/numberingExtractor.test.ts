/**
 * @file numberingExtractor.test.ts
 * Unit tests for OOXML numbering definition extraction.
 */

import { extractNumberingMaps } from '../extractor/numberingExtractor.js';

describe('numberingExtractor', () => {
  it('should detect bullet list type (numFmt="bullet" maps to "bullet")', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="bullet"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="10">
        <w:abstractNumId w:val="1"/>
      </w:num>
    </w:numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(abstractNumMap.get('1:0')).toBe('bullet');
    expect(numIdMap.get(10)).toBe(1);
  });

  it('should detect ordered list type (numFmt="decimal" maps to "ordered")', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="2">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="decimal"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="20">
        <w:abstractNumId w:val="2"/>
      </w:num>
    </w:numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(abstractNumMap.get('2:0')).toBe('ordered');
    expect(numIdMap.get(20)).toBe(2);
  });

  it('should handle multiple levels (ilvl 0, 1, 2) in the same abstractNum', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="5">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="decimal"/>
        </w:lvl>
        <w:lvl w:ilvl="1">
          <w:numFmt w:val="bullet"/>
        </w:lvl>
        <w:lvl w:ilvl="2">
          <w:numFmt w:val="decimal"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1">
        <w:abstractNumId w:val="5"/>
      </w:num>
    </w:numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(abstractNumMap.get('5:0')).toBe('ordered');
    expect(abstractNumMap.get('5:1')).toBe('bullet');
    expect(abstractNumMap.get('5:2')).toBe('ordered');
    expect(numIdMap.get(1)).toBe(5);
  });

  it('should correctly map numId to abstractNumId with multiple w:num elements', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="bullet"/>
        </w:lvl>
      </w:abstractNum>
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="decimal"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="100">
        <w:abstractNumId w:val="0"/>
      </w:num>
      <w:num w:numId="200">
        <w:abstractNumId w:val="1"/>
      </w:num>
      <w:num w:numId="300">
        <w:abstractNumId w:val="0"/>
      </w:num>
    </w:numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(numIdMap.get(100)).toBe(0);
    expect(numIdMap.get(200)).toBe(1);
    expect(numIdMap.get(300)).toBe(0);
    expect(abstractNumMap.get('0:0')).toBe('bullet');
    expect(abstractNumMap.get('1:0')).toBe('ordered');
  });

  it('should return empty maps without throwing when given an empty string', () => {
    const result = extractNumberingMaps('');
    expect(result.numIdMap.size).toBe(0);
    expect(result.abstractNumMap.size).toBe(0);
  });

  it('should return empty maps without throwing when given whitespace or invalid/malformed XML', () => {
    const wsResult = extractNumberingMaps('   \n  \t  ');
    expect(wsResult.numIdMap.size).toBe(0);
    expect(wsResult.abstractNumMap.size).toBe(0);

    const malformedResult = extractNumberingMaps('<not valid <xml');
    expect(malformedResult.numIdMap.size).toBe(0);
    expect(malformedResult.abstractNumMap.size).toBe(0);
  });

  it('should map lowerRoman and upperLetter formats to "ordered"', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="3">
        <w:lvl w:ilvl="0">
          <w:numFmt w:val="lowerRoman"/>
        </w:lvl>
        <w:lvl w:ilvl="1">
          <w:numFmt w:val="upperLetter"/>
        </w:lvl>
        <w:lvl w:ilvl="2">
          <w:numFmt w:val="lowerLetter"/>
        </w:lvl>
        <w:lvl w:ilvl="3">
          <w:numFmt w:val="upperRoman"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="3">
        <w:abstractNumId w:val="3"/>
      </w:num>
    </w:numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(abstractNumMap.get('3:0')).toBe('ordered');
    expect(abstractNumMap.get('3:1')).toBe('ordered');
    expect(abstractNumMap.get('3:2')).toBe('ordered');
    expect(abstractNumMap.get('3:3')).toBe('ordered');
    expect(numIdMap.get(3)).toBe(3);
  });

  it('should handle XML without namespace prefixes', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <numbering>
      <abstractNum abstractNumId="7">
        <lvl ilvl="0">
          <numFmt val="bullet"/>
        </lvl>
      </abstractNum>
      <num numId="77">
        <abstractNumId val="7"/>
      </num>
    </numbering>`;

    const { numIdMap, abstractNumMap } = extractNumberingMaps(xml);
    expect(abstractNumMap.get('7:0')).toBe('bullet');
    expect(numIdMap.get(77)).toBe(7);
  });
});
