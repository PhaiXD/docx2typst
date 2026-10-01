/**
 * @file tableExtractor.test.ts
 * Unit tests for table extraction from OOXML markup.
 */

import { extractTables } from '../extractor/tableExtractor.js';
import { parseXml } from '../utils/xmlParser.js';

/**
 * Helper to parse an XML string and retrieve the body node.
 *
 * @param xml - OOXML XML string.
 * @returns The parsed w:body record.
 */
function getBodyNode(xml: string): Record<string, unknown> {
  const parsed = parseXml(xml);
  const docRoot = (parsed['w:document'] || parsed['document'] || parsed) as Record<string, unknown>;
  const body = (docRoot['w:body'] || docRoot['body'] || docRoot) as Record<string, unknown>;
  return body;
}

describe('tableExtractor', () => {
  it('should extract a simple 2x2 table (2 rows, 2 columns, correct cell content)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="TableGrid"/>
          </w:tblPr>
          <w:tr>
            <w:tc>
              <w:p>
                <w:r><w:t>R1C1</w:t></w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:p>
                <w:r><w:t>R1C2</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:p>
                <w:r><w:t>R2C1</w:t></w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:p>
                <w:r><w:t>R2C2</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    const table = tables[0];
    expect(table.columnCount).toBe(2);
    expect(table.style).toBe('TableGrid');
    expect(table.rows).toHaveLength(2);

    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[0].cells[0].paragraphs[0].text).toBe('R1C1');
    expect(table.rows[0].cells[1].paragraphs[0].text).toBe('R1C2');

    expect(table.rows[1].cells).toHaveLength(2);
    expect(table.rows[1].cells[0].paragraphs[0].text).toBe('R2C1');
    expect(table.rows[1].cells[1].paragraphs[0].text).toBe('R2C2');
  });

  it('should detect table column count from first row', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc><w:p><w:r><w:t>Col 1</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Col 2</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Col 3</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].columnCount).toBe(3);
  });

  it('should detect column count from tblGrid when first row has no cells', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tblGrid>
            <w:gridCol/>
            <w:gridCol/>
            <w:gridCol/>
            <w:gridCol/>
          </w:tblGrid>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].columnCount).toBe(4);
  });

  it('should extract cell with columnSpan > 1 (gridSpan)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:gridSpan w:val="3"/>
              </w:tcPr>
              <w:p><w:r><w:t>Spanning 3 columns</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].rows[0].cells[0].columnSpan).toBe(3);
  });

  it('should detect header row when paragraph style contains Heading1 in first row', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:p>
                <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                <w:r><w:t>Header A</w:t></w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:p>
                <w:r><w:t>Header B</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:p><w:r><w:t>Data A</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:p><w:r><w:t>Data B</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].rows[0].isHeader).toBe(true);
    expect(tables[0].rows[1].isHeader).toBeUndefined();
  });

  it('should not detect header when no heading style (regular style or no style)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:p>
                <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                <w:r><w:t>Regular Cell 1</w:t></w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:p>
                <w:r><w:t>Regular Cell 2</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:p><w:r><w:t>Row 2 Cell 1</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
              <w:p><w:r><w:t>Row 2 Cell 2</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].rows[0].isHeader).toBeUndefined();
    expect(tables[0].rows[1].isHeader).toBeUndefined();
  });

  it('should extract cell with multiple paragraphs (content joined with newline)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:p><w:r><w:t>First line of cell</w:t></w:r></w:p>
              <w:p><w:r><w:t>Second line of cell</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    const cell = tables[0].rows[0].cells[0];
    expect(cell.paragraphs).toHaveLength(2);
    expect(cell.paragraphs[0].text).toBe('First line of cell');
    expect(cell.paragraphs[1].text).toBe('Second line of cell');
  });

  it('should return empty array for empty table body or non-table content', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>No table here</w:t></w:r></w:p>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    expect(extractTables(body)).toEqual([]);
    expect(extractTables({} as Record<string, unknown>)).toEqual([]);
    expect(extractTables(null as unknown as Record<string, unknown>)).toEqual([]);
  });

  it('should extract vertical merge indicator (w:vMerge)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:vMerge w:val="restart"/>
              </w:tcPr>
              <w:p><w:r><w:t>Restart</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:vMerge/>
              </w:tcPr>
              <w:p><w:r><w:t>Continuation</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);

    expect(tables).toHaveLength(1);
    expect(tables[0].rows[0].cells[0].isVerticalMerge).toBe(true);
    expect(tables[0].rows[1].cells[0].isVerticalMerge).toBe(true);
  });

  it('should extract table style when provided as a string node', () => {
    const body = {
      'w:tbl': {
        'w:tblPr': {
          'w:tblStyle': 'CustomTableStyle',
        },
        'w:tr': {
          'w:tc': {
            'w:p': {
              'w:r': { 'w:t': 'Cell' },
            },
          },
        },
      },
    };

    const tables = extractTables(body as unknown as Record<string, unknown>);
    expect(tables).toHaveLength(1);
    expect(tables[0].style).toBe('CustomTableStyle');
  });

  it('should fall back to counting cells in subsequent rows if first row is empty', () => {
    const body = {
      'w:tbl': {
        'w:tr': [
          {},
          {
            'w:tc': [
              { 'w:p': { 'w:r': { 'w:t': '1' } } },
              { 'w:p': { 'w:r': { 'w:t': '2' } } },
            ],
          },
        ],
      },
    };

    const tables = extractTables(body as unknown as Record<string, unknown>);
    expect(tables).toHaveLength(1);
    expect(tables[0].columnCount).toBe(2);
  });

  it('should handle gridSpan as number or numeric string directly', () => {
    const body = {
      'w:tbl': {
        'w:tr': {
          'w:tc': [
            {
              'w:tcPr': { 'w:gridSpan': 2 },
              'w:p': { 'w:r': { 'w:t': 'Span 2' } },
            },
            {
              'w:tcPr': { 'w:gridSpan': '3' },
              'w:p': { 'w:r': { 'w:t': 'Span 3' } },
            },
          ],
        },
      },
    };

    const tables = extractTables(body as unknown as Record<string, unknown>);
    expect(tables).toHaveLength(1);
    expect(tables[0].rows[0].cells[0].columnSpan).toBe(2);
    expect(tables[0].rows[0].cells[1].columnSpan).toBe(3);
  });

  it('should extract table borders from w:tblBorders', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tblPr>
            <w:tblBorders>
              <w:top w:val="single"/>
              <w:bottom w:val="single"/>
              <w:left w:val="none"/>
              <w:right w:val="none"/>
              <w:insideH w:val="single"/>
              <w:insideV w:val="none"/>
            </w:tblBorders>
          </w:tblPr>
          <w:tr>
            <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);
    expect(tables[0].borders).toEqual({
      top: true,
      bottom: true,
      left: false,
      right: false,
      insideH: true,
      insideV: false,
    });
  });

  it('should clear vertical borders when cells explicitly set left/right to nil', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tblPr>
            <w:tblBorders>
              <w:top w:val="single"/>
              <w:bottom w:val="single"/>
              <w:left w:val="single"/>
              <w:right w:val="single"/>
              <w:insideH w:val="single"/>
              <w:insideV w:val="single"/>
            </w:tblBorders>
          </w:tblPr>
          <w:tr>
            <w:tc>
              <w:tcPr>
                <w:tcBorders>
                  <w:top w:val="single"/>
                  <w:bottom w:val="single"/>
                  <w:left w:val="nil"/>
                  <w:right w:val="nil"/>
                </w:tcBorders>
              </w:tcPr>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>`;

    const body = getBodyNode(xml);
    const tables = extractTables(body);
    expect(tables[0].borders).toEqual({
      top: true,
      bottom: true,
      left: false,
      right: false,
      insideH: true,
      insideV: false,
    });
  });
});

