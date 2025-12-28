/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect, afterEach } from 'bun:test';
import { cell } from '@sheet/cell';
import { parseSheet } from '@sheet/reader';
import { row } from '@sheet/row';
import { parseXmlEvents } from '@xml/parser';
import { readXlsx } from './reader';
import { writeXlsx } from './writer';

describe('XLSXReader', () => {
  const testFile = 'test-read.xlsx';

  afterEach(async () => {
    if (await Bun.file(testFile).exists()) {
      await import('fs').then((fs) => fs.promises.unlink(testFile));
    }
  });

  test('should read single sheet workbook', async () => {
    // Write a test file first
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            yield row([cell('Name'), cell('Age')]);
            yield row([cell('Alice'), cell(30)]);
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(1);

    const sheet = workbook.sheet('Data');
    expect(sheet.name).toBe('Data');

    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('Name');
    expect(rows[1]?.cells[0]?.value).toBe('Alice');
  });

  test('should read multiple sheets workbook', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'Sheet1', rows: (async function* () { yield row([cell('A')]); })() },
        { name: 'Sheet2', rows: (async function* () { yield row([cell('B')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(2);

    const sheet1 = workbook.sheet('Sheet1');
    expect(sheet1.name).toBe('Sheet1');

    const sheet2 = workbook.sheet(1);
    expect(sheet2.name).toBe('Sheet2');
  });

  test('should get sheet by name', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'MySheet', rows: (async function* () { yield row([cell('Test')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('MySheet');
    expect(sheet.name).toBe('MySheet');
  });

  test('should get sheet by index', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'First', rows: (async function* () { yield row([cell('A')]); })() },
        { name: 'Second', rows: (async function* () { yield row([cell('B')]); })() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheet0 = workbook.sheet(0);
    expect(sheet0.name).toBe('First');

    const sheet1 = workbook.sheet(1);
    expect(sheet1.name).toBe('Second');
  });

  test('should iterate all sheets', async () => {
    await writeXlsx(testFile, {
      sheets: [
        { name: 'A', rows: (async function* () {})() },
        { name: 'B', rows: (async function* () {})() },
        { name: 'C', rows: (async function* () {})() },
      ],
    });

    const workbook = await readXlsx(testFile);
    const sheets = workbook.sheets();
    expect(sheets).toHaveLength(3);
    expect(sheets[0]?.name).toBe('A');
    expect(sheets[1]?.name).toBe('B');
    expect(sheets[2]?.name).toBe('C');
  });

  test('should handle error for invalid files', async () => {
    await expect(readXlsx('nonexistent.xlsx')).rejects.toThrow();
  });

  test('should handle error for missing sheets', async () => {
    await writeXlsx(testFile, {
      sheets: [{ name: 'Data', rows: (async function* () {})() }],
    });

    const workbook = await readXlsx(testFile);
    expect(() => workbook.sheet('NonExistent')).toThrow();
  });

  test('should read multiple sheets and verify isolation', async () => {
    // Write workbook with multiple sheets, each with different data
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Sheet1',
          rows: (async function* () {
            yield row([cell('Sheet1-Row1'), cell(1)]);
            yield row([cell('Sheet1-Row2'), cell(2)]);
          })(),
        },
        {
          name: 'Sheet2',
          rows: (async function* () {
            yield row([cell('Sheet2-Row1'), cell(10)]);
            yield row([cell('Sheet2-Row2'), cell(20)]);
            yield row([cell('Sheet2-Row3'), cell(30)]);
          })(),
        },
        {
          name: 'Sheet3',
          rows: (async function* () {
            yield row([cell('Sheet3-Only')]);
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(3);

    // Read Sheet1 and verify it only contains Sheet1 data
    const sheet1 = workbook.sheet('Sheet1');
    const sheet1Rows: any[] = [];
    for await (const row of sheet1.rows()) {
      sheet1Rows.push(row);
    }
    expect(sheet1Rows).toHaveLength(2);
    expect(sheet1Rows[0]?.cells[0]?.value).toBe('Sheet1-Row1');
    expect(sheet1Rows[0]?.cells[1]?.value).toBe(1);
    expect(sheet1Rows[1]?.cells[0]?.value).toBe('Sheet1-Row2');
    expect(sheet1Rows[1]?.cells[1]?.value).toBe(2);
    // Verify no Sheet2 or Sheet3 data
    expect(sheet1Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet1'))).toBe(true);

    // Read Sheet2 and verify it only contains Sheet2 data
    const sheet2 = workbook.sheet('Sheet2');
    const sheet2Rows: any[] = [];
    for await (const row of sheet2.rows()) {
      sheet2Rows.push(row);
    }
    expect(sheet2Rows).toHaveLength(3);
    expect(sheet2Rows[0]?.cells[0]?.value).toBe('Sheet2-Row1');
    expect(sheet2Rows[0]?.cells[1]?.value).toBe(10);
    expect(sheet2Rows[1]?.cells[0]?.value).toBe('Sheet2-Row2');
    expect(sheet2Rows[1]?.cells[1]?.value).toBe(20);
    expect(sheet2Rows[2]?.cells[0]?.value).toBe('Sheet2-Row3');
    expect(sheet2Rows[2]?.cells[1]?.value).toBe(30);
    // Verify no Sheet1 or Sheet3 data
    expect(sheet2Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet2'))).toBe(true);

    // Read Sheet3 and verify it only contains Sheet3 data
    const sheet3 = workbook.sheet('Sheet3');
    const sheet3Rows: any[] = [];
    for await (const row of sheet3.rows()) {
      sheet3Rows.push(row);
    }
    expect(sheet3Rows).toHaveLength(1);
    expect(sheet3Rows[0]?.cells[0]?.value).toBe('Sheet3-Only');
    // Verify no Sheet1 or Sheet2 data
    expect(sheet3Rows.every((r) => r.cells[0]?.value?.toString().startsWith('Sheet3'))).toBe(true);
  });

  test('should read workbook with shared strings', async () => {
    // Write a file with shared strings
    await writeXlsx(
      testFile,
      {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Hello'), cell('World')]);
              yield row([cell('Hello'), cell('Test')]); // Duplicate string
            })(),
          },
        ],
      },
      { sharedStrings: 'shared' },
    );

    // Read it back
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    // Verify strings are correctly resolved from shared strings table
    expect(rows).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('Hello');
    expect(rows[0]?.cells[1]?.value).toBe('World');
    expect(rows[1]?.cells[0]?.value).toBe('Hello'); // Should resolve from shared strings
    expect(rows[1]?.cells[1]?.value).toBe('Test');
  });

  test('should preserve newlines in multiline strings', async () => {
    // Create a file with multiline strings containing newlines
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Multiline',
          rows: (async function* () {
            yield row([cell('Line 1\nLine 2\nLine 3')]);
            yield row([cell('No newlines here')]);
            yield row([cell('Before\n\nAfter double newline')]);
          })(),
        },
      ],
    });

    // Read back and verify newlines are preserved
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Multiline');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('Line 1\nLine 2\nLine 3');
    expect(rows[1]?.cells[0]?.value).toBe('No newlines here');
    expect(rows[2]?.cells[0]?.value).toBe('Before\n\nAfter double newline');

    // Verify the strings contain actual newline characters, not escaped versions
    expect(rows[0]?.cells[0]?.value.split('\n')).toHaveLength(3);
    expect(rows[2]?.cells[0]?.value.split('\n')).toHaveLength(3); // Before, empty line, After
  });

  test('should handle XML entities in strings', async () => {
    // Create a file with strings containing XML entities
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Entities',
          rows: (async function* () {
            yield row([cell('<test> & "quotes" \'single\'')]);
            yield row([cell('Normal text')]);
            yield row([cell('© ® ™ € £ ¥')]);
          })(),
        },
      ],
    });

    // Read back and verify XML entities are properly decoded
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Entities');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('<test> & "quotes" \'single\'');
    expect(rows[1]?.cells[0]?.value).toBe('Normal text');
    expect(rows[2]?.cells[0]?.value).toBe('© ® ™ € £ ¥');
  });

  test('should skip pronunciation data in Japanese text', async () => {
    // Create XML with Japanese text containing pronunciation data (ruby annotations)
    // This simulates what Excel might generate for Japanese text with furigana
    const xmlWithRuby = `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="str">
        <is>
          <t>東京</t>
          <rPh sb="0" eb="1">
            <t>とう</t>
          </rPh>
          <rPh sb="1" eb="2">
            <t>きょう</t>
          </rPh>
        </is>
      </c>
    </row>
    <row r="2">
      <c r="A2" t="str">
        <is>
          <t>日本語</t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithRuby);
    }();

    // Parse the sheet directly to test pronunciation filtering
    const rows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(2);
    // Should return only the main text, filtering out pronunciation data
    expect(rows[0]?.cells[0]?.value).toBe('東京'); // Main text only, no furigana
    expect(rows[1]?.cells[0]?.value).toBe('日本語'); // Normal text unchanged
  });

  test('should identify hidden sheets correctly', async () => {
    // Create a file with both visible and hidden sheets
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Visible Sheet',
          rows: (async function* () {
            yield row([cell('Visible Data')]);
          })(),
        },
        {
          name: 'Hidden Sheet',
          hidden: true,
          rows: (async function* () {
            yield row([cell('Hidden Data')]);
          })(),
        },
        {
          name: 'Another Visible',
          rows: (async function* () {
            yield row([cell('More Data')]);
          })(),
        },
      ],
    });

    // Read the file and check sheet visibility
    const workbook = await readXlsx(testFile);
    const sheets = workbook.sheets();

    expect(sheets).toHaveLength(3);

    // Check visible sheets
    const visibleSheet1 = workbook.sheet('Visible Sheet');
    expect(visibleSheet1.hidden).toBe(false);
    expect(visibleSheet1.name).toBe('Visible Sheet');

    const visibleSheet2 = workbook.sheet('Another Visible');
    expect(visibleSheet2.hidden).toBe(false);
    expect(visibleSheet2.name).toBe('Another Visible');

    // Check hidden sheet
    const hiddenSheet = workbook.sheet('Hidden Sheet');
    expect(hiddenSheet.hidden).toBe(true);
    expect(hiddenSheet.name).toBe('Hidden Sheet');

    // Verify all sheets are accessible regardless of visibility
    const allSheets = workbook.sheets();
    const hiddenSheets = allSheets.filter((s) => s.hidden);
    const visibleSheets = allSheets.filter((s) => !s.hidden);

    expect(hiddenSheets).toHaveLength(1);
    expect(hiddenSheets[0]?.name).toBe('Hidden Sheet');
    expect(visibleSheets).toHaveLength(2);
    expect(visibleSheets.map((s) => s.name)).toContain('Visible Sheet');
    expect(visibleSheets.map((s) => s.name)).toContain('Another Visible');
  });

  test('should preserve empty cells at row ends when spans are present', async () => {
    // Create a file with rows that have trailing empty cells
    // Row 1: Data in A-C, empty D-E (spans="1:5" means 5 columns)
    // Row 2: Data in A-B, empty C-E (spans="1:5" means 5 columns)
    // Row 3: No spans, only data in A-B (should not be padded)
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Spans',
          rows: (async function* () {
            // Row with data in first 3 columns, should span to 5 columns
            yield row([cell('A1'), cell('B1'), cell('C1')], { rowIndex: 1 });
            // Row with data in first 2 columns, should span to 5 columns
            yield row([cell('A2'), cell('B2')], { rowIndex: 2 });
            // Row without explicit span requirement (no trailing padding expected)
            yield row([cell('A3'), cell('B3')], { rowIndex: 3 });
          })(),
        },
      ],
    });

    // Read the file back
    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Spans');
    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3);

    // Row 1: Should have 5 cells (A-C with data, D-E empty) when spans indicate 5 columns
    // Note: The writer generates spans based on actual cell count, so we need to test with
    // a manually created XML that has explicit spans
    // For now, verify the basic structure works

    // Create a test with explicit XML that includes spans
    const xmlWithSpans = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1" spans="1:5">
      <c r="A1" t="inlineStr"><is><t>ColA</t></is></c>
      <c r="B1" t="inlineStr"><is><t>ColB</t></is></c>
      <c r="C1" t="inlineStr"><is><t>ColC</t></is></c>
    </row>
    <row r="2" spans="1:5">
      <c r="A2" t="inlineStr"><is><t>DataA</t></is></c>
      <c r="B2" t="inlineStr"><is><t>DataB</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>OnlyA</t></is></c>
      <c r="B3" t="inlineStr"><is><t>OnlyB</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithSpans);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      parsedRows.push(row);
    }

    expect(parsedRows).toHaveLength(3);

    // Row 1: spans="1:5" means 5 columns (indices 0-4)
    // Has cells A1, B1, C1, so should have 5 cells total with D1 and E1 empty
    expect(parsedRows[0]?.cells).toHaveLength(5);
    expect(parsedRows[0]?.cells[0]?.value).toBe('ColA'); // A1
    expect(parsedRows[0]?.cells[1]?.value).toBe('ColB'); // B1
    expect(parsedRows[0]?.cells[2]?.value).toBe('ColC'); // C1
    expect(parsedRows[0]?.cells[3]?.value).toBe(''); // D1 (empty, padded)
    expect(parsedRows[0]?.cells[4]?.value).toBe(''); // E1 (empty, padded)

    // Row 2: spans="1:5" means 5 columns
    // Has cells A2, B2, so should have 5 cells total with C2, D2, E2 empty
    expect(parsedRows[1]?.cells).toHaveLength(5);
    expect(parsedRows[1]?.cells[0]?.value).toBe('DataA'); // A2
    expect(parsedRows[1]?.cells[1]?.value).toBe('DataB'); // B2
    expect(parsedRows[1]?.cells[2]?.value).toBe(''); // C2 (empty, padded)
    expect(parsedRows[1]?.cells[3]?.value).toBe(''); // D2 (empty, padded)
    expect(parsedRows[1]?.cells[4]?.value).toBe(''); // E2 (empty, padded)

    // Row 3: No spans attribute, should only have cells that are present
    expect(parsedRows[2]?.cells).toHaveLength(2);
    expect(parsedRows[2]?.cells[0]?.value).toBe('OnlyA'); // A3
    expect(parsedRows[2]?.cells[1]?.value).toBe('OnlyB'); // B3
  });

  test('should read cells without cell references', async () => {
    // Create XML with cells that lack explicit r="A1" attributes
    // Cells should be read correctly by their position in the row
    const xmlWithoutRefs = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c t="inlineStr"><is><t>First</t></is></c>
      <c t="inlineStr"><is><t>Second</t></is></c>
      <c><v>42</v></c>
      <c t="inlineStr"><is><t>Fourth</t></is></c>
    </row>
    <row r="2">
      <c><v>100</v></c>
      <c><v>200</v></c>
      <c><v>300</v></c>
    </row>
    <row r="3">
      <c t="inlineStr"><is><t>Only</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithoutRefs);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      parsedRows.push(row);
    }

    expect(parsedRows).toHaveLength(3);

    // Row 1: 4 cells without references, should be read by position
    expect(parsedRows[0]?.cells).toHaveLength(4);
    expect(parsedRows[0]?.cells[0]?.value).toBe('First');
    expect(parsedRows[0]?.cells[0]?.type).toBe('string');
    expect(parsedRows[0]?.cells[1]?.value).toBe('Second');
    expect(parsedRows[0]?.cells[1]?.type).toBe('string');
    expect(parsedRows[0]?.cells[2]?.value).toBe(42);
    expect(parsedRows[0]?.cells[2]?.type).toBe('number');
    expect(parsedRows[0]?.cells[3]?.value).toBe('Fourth');
    expect(parsedRows[0]?.cells[3]?.type).toBe('string');

    // Row 2: 3 numeric cells without references
    expect(parsedRows[1]?.cells).toHaveLength(3);
    expect(parsedRows[1]?.cells[0]?.value).toBe(100);
    expect(parsedRows[1]?.cells[1]?.value).toBe(200);
    expect(parsedRows[1]?.cells[2]?.value).toBe(300);

    // Row 3: Single cell without reference
    expect(parsedRows[2]?.cells).toHaveLength(1);
    expect(parsedRows[2]?.cells[0]?.value).toBe('Only');
    expect(parsedRows[2]?.cells[0]?.type).toBe('string');
  });

  test('should handle invalid dates', async () => {
    // Create XML with date cells that have invalid timestamp values
    // Invalid dates include: out-of-range values, NaN, Infinity
    const xmlWithInvalidDates = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c t="d"><v>1</v></c>
      <c t="d"><v>-700000</v></c>
      <c t="d"><v>3000000</v></c>
      <c t="d"><v>44239</v></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithInvalidDates);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { use1904Dates: false })) {
      parsedRows.push(row);
    }

    expect(parsedRows).toHaveLength(1);
    expect(parsedRows[0]?.cells).toHaveLength(4);

    // Valid date (Excel day 1 = January 1, 1900)
    expect(parsedRows[0]?.cells[0]?.type).toBe('date');
    expect(parsedRows[0]?.cells[0]?.value).toBeInstanceOf(Date);
    expect((parsedRows[0]?.cells[0]?.value as Date)?.getFullYear()).toBe(1900);

    // Invalid date: out of range (too negative)
    expect(parsedRows[0]?.cells[1]?.type).toBe('date');
    expect(parsedRows[0]?.cells[1]?.value).toBeNull();

    // Invalid date: out of range (too large)
    expect(parsedRows[0]?.cells[2]?.type).toBe('date');
    expect(parsedRows[0]?.cells[2]?.value).toBeNull();

    // Valid date (Excel day 44239 = January 1, 2021)
    expect(parsedRows[0]?.cells[3]?.type).toBe('date');
    expect(parsedRows[0]?.cells[3]?.value).toBeInstanceOf(Date);
    expect((parsedRows[0]?.cells[3]?.value as Date)?.getFullYear()).toBe(2021);
  });

  test('should handle empty sheets', async () => {
    // Create a file with both empty and non-empty sheets
    await writeXlsx(testFile, {
      sheets: [
        {
          name: 'Empty Sheet 1',
          rows: (async function* () {
            // Empty sheet - no rows yielded
          })(),
        },
        {
          name: 'Data Sheet',
          rows: (async function* () {
            yield row([cell('Data')]);
          })(),
        },
        {
          name: 'Empty Sheet 2',
          rows: (async function* () {
            // Empty sheet - no rows yielded
          })(),
        },
      ],
    });

    const workbook = await readXlsx(testFile);
    expect(workbook.sheets()).toHaveLength(3);

    // Test first empty sheet
    const emptySheet1 = workbook.sheet('Empty Sheet 1');
    expect(emptySheet1.name).toBe('Empty Sheet 1');
    const rows1: any[] = [];
    for await (const row of emptySheet1.rows()) {
      rows1.push(row);
    }
    expect(rows1).toHaveLength(0);

    // Test data sheet (should have rows)
    const dataSheet = workbook.sheet('Data Sheet');
    expect(dataSheet.name).toBe('Data Sheet');
    const dataRows: any[] = [];
    for await (const row of dataSheet.rows()) {
      dataRows.push(row);
    }
    expect(dataRows).toHaveLength(1);
    expect(dataRows[0]?.cells[0]?.value).toBe('Data');

    // Test second empty sheet
    const emptySheet2 = workbook.sheet('Empty Sheet 2');
    expect(emptySheet2.name).toBe('Empty Sheet 2');
    const rows2: any[] = [];
    for await (const row of emptySheet2.rows()) {
      rows2.push(row);
    }
    expect(rows2).toHaveLength(0);

    // Test accessing empty sheet by index
    const emptySheetByIndex = workbook.sheet(0);
    expect(emptySheetByIndex.name).toBe('Empty Sheet 1');
    const rowsByIndex: any[] = [];
    for await (const row of emptySheetByIndex.rows()) {
      rowsByIndex.push(row);
    }
    expect(rowsByIndex).toHaveLength(0);
  });

  test('should handle prefixed XML files', async () => {
    // Test that XML with namespace prefixes (like dc:title, cp:keywords) is parsed correctly
    // This simulates reading core.xml or other files with namespace prefixes
    const xmlWithPrefixes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <cp:keywords>test, xml, prefixes</cp:keywords>
</cp:coreProperties>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithPrefixes);
    }();

    const events: any[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    // Verify that prefixed elements are parsed correctly
    const corePropsStart = events.find(
      (e) => e.type === 'startElement' && e.name === 'cp:coreProperties',
    );
    expect(corePropsStart).toBeDefined();
    expect(corePropsStart?.attributes).toBeDefined();
    expect(corePropsStart?.attributes['xmlns:cp']).toBe(
      'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    );
    expect(corePropsStart?.attributes['xmlns:dc']).toBe('http://purl.org/dc/elements/1.1/');

    const titleStart = events.find((e) => e.type === 'startElement' && e.name === 'dc:title');
    expect(titleStart).toBeDefined();

    const titleText = events.find(
      (e) => e.type === 'text' && events.indexOf(e) > events.indexOf(titleStart!),
    );
    expect(titleText).toBeDefined();
    expect(titleText?.text).toBe('Test Document');

    const creatorStart = events.find(
      (e) => e.type === 'startElement' && e.name === 'dc:creator',
    );
    expect(creatorStart).toBeDefined();

    const keywordsStart = events.find(
      (e) => e.type === 'startElement' && e.name === 'cp:keywords',
    );
    expect(keywordsStart).toBeDefined();
  });

  test('should handle strict OOXML files', async () => {
    // Test that strict OOXML format (with proper namespace declarations) works correctly
    // Strict OOXML requires all namespaces to be properly declared
    const strictOOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <c r="A1" t="inlineStr" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <is xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <t xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">Strict OOXML</t>
        </is>
      </c>
      <c r="B1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">42</v>
      </c>
    </row>
    <row r="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <c r="A2" t="s" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">0</v>
      </c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(strictOOXML);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      parsedRows.push(row);
    }

    // Verify that strict OOXML format is parsed correctly
    expect(parsedRows).toHaveLength(2);

    // First row: inline string and number
    expect(parsedRows[0]?.cells).toHaveLength(2);
    expect(parsedRows[0]?.cells[0]?.value).toBe('Strict OOXML');
    expect(parsedRows[0]?.cells[0]?.type).toBe('string');
    expect(parsedRows[0]?.cells[1]?.value).toBe(42);
    expect(parsedRows[0]?.cells[1]?.type).toBe('number'); // Number type is inferred

    // Second row: shared string reference
    expect(parsedRows[1]?.cells).toHaveLength(1);
    // Note: shared string index 0 would need to be resolved, but we're testing parsing
    expect(parsedRows[1]?.cells[0]?.type).toBe('string');
  });

  test('should handle missing shared strings count metadata', async () => {
    // Test that shared strings can be parsed even when count/uniqueCount attributes are missing
    // This simulates files with incomplete metadata
    const xmlWithoutCount = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>First</t></si>
  <si><t>Second</t></si>
  <si><t>Third</t></si>
</sst>`;

    // Import the parser directly
    const { parseSharedStrings } = await import('./shared-strings-reader');
    const { openZip } = await import('@zip/reader');
    const { createZipWriter, writeZipEntry, endZipWriter } = await import('@zip/writer');
    const { stringToBytes } = await import('../adapters/common');

    // Create a temporary ZIP with shared strings file
    const zipWriter = createZipWriter();
    await writeZipEntry(zipWriter, 'xl/sharedStrings.xml', stringToBytes(xmlWithoutCount));
    const buffer = await endZipWriter(zipWriter);

    const zipFile = await openZip(buffer);
    const sharedStringsEntry = zipFile.entries.find(
      (e) => e.fileName === 'xl/sharedStrings.xml',
    );
    expect(sharedStringsEntry).toBeDefined();

    const strategy = await parseSharedStrings(sharedStringsEntry!, zipFile.zipFile);
    // Should parse all strings even without count metadata
    expect(strategy.getCount()).toBe(3);
    expect(await strategy.getString(0)).toBe('First');
    expect(await strategy.getString(1)).toBe('Second');
    expect(await strategy.getString(2)).toBe('Third');

    // Cleanup
    await strategy.cleanup();
  });

  test('should handle capital shared strings filename', async () => {
    // Test that shared strings can be found even with different case filename
    // Some tools may create files with different casing
    await writeXlsx(
      testFile,
      {
        sheets: [
          {
            name: 'Data',
            rows: (async function* () {
              yield row([cell('Test')]);
            })(),
          },
        ],
      },
      { sharedStrings: 'shared' },
    );

    // Manually rename the shared strings file to capital case
    const { openZip } = await import('@zip/reader');
    const { createZipWriter, writeZipEntry, endZipWriter } = await import('@zip/writer');
    const { readZipEntry } = await import('@zip/reader');
    const { stringToBytes, bytesToString } = await import('../adapters/common');

    const file = Bun.file(testFile);
    const buffer = Buffer.from(await file.arrayBuffer());
    const originalZip = await openZip(buffer);

    // Find and read the original shared strings file
    const originalEntry = originalZip.entries.find(
      (e) => e.fileName === 'xl/sharedStrings.xml',
    );
    expect(originalEntry).toBeDefined();

    // Read the content
    const contentBytes: Uint8Array[] = [];
    for await (const chunk of readZipEntry(originalEntry!, originalZip.zipFile)) {
      contentBytes.push(chunk);
    }
    const content = await bytesToString(async function* () {
      for (const chunk of contentBytes) yield chunk;
    }());

    // Create new ZIP with capital case filename
    const newZip = createZipWriter();
    // Copy all entries except the shared strings one
    for (const entry of originalZip.entries) {
      if (entry.fileName !== 'xl/sharedStrings.xml') {
        const entryBytes: Uint8Array[] = [];
        for await (const chunk of readZipEntry(entry, originalZip.zipFile)) {
          entryBytes.push(chunk);
        }
        const entryContent = await bytesToString(async function* () {
          for (const chunk of entryBytes) yield chunk;
        }());
        await writeZipEntry(newZip, entry.fileName, stringToBytes(entryContent));
      }
    }
    // Add shared strings with capital case filename
    await writeZipEntry(newZip, 'xl/SharedStrings.xml', stringToBytes(content));
    const newBuffer = await endZipWriter(newZip);
    await Bun.write(testFile, newBuffer);

    const workbook = await readXlsx(testFile);
    const sheet = workbook.sheet('Data');

    const rows: any[] = [];
    for await (const row of sheet.rows()) {
      rows.push(row);
    }

    // Should find shared strings even with capital case filename
    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells[0]?.value).toBe('Test'); // Should resolve from shared strings
    expect(rows[0]?.cells[0]?.type).toBe('string');
  });
});

