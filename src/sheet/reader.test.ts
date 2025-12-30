/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, test, expect } from 'bun:test';
import { parseXmlEvents } from '@xml/parser';
import { parseSheet } from './reader';
import type { Row } from '../types';

describe('Row Parser', () => {
  test('should parse single row with one cell', async () => {
    const xml = '<row><c><v>Hello</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    expect(rows[0]?.cells[0]?.value).toBe('Hello');
  });

  test('should parse row with multiple cells', async () => {
    const xml = '<row><c><v>A</v></c><c><v>1</v></c><c><v>B</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('A');
    expect(rows[0]?.cells[1]?.value).toBe(1);
    expect(rows[0]?.cells[2]?.value).toBe('B');
  });

  test('should parse row with cell types (string)', async () => {
    const xml = '<row><c t="s"><v>Hello</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows[0]?.cells[0]?.type).toBe('string');
    expect(rows[0]?.cells[0]?.value).toBe('Hello');
  });

  test('should parse row with cell types (number)', async () => {
    const xml = '<row><c><v>42</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.value).toBe(42);
    expect(cell?.type).toBe('number');
  });

  test('should parse formula cell with numeric computed value', async () => {
    const xml = '<row><c><f>SUM(A1:A5)</f><v>42</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=SUM(A1:A5)');
    expect(cell?.formula).toBe('SUM(A1:A5)');
    expect(cell?.computedValue).toBe(42);
  });

  test('should parse formula cell with string computed value', async () => {
    const xml = '<row><c><f>CONCATENATE(A1,B1)</f><v>Hello World</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=CONCATENATE(A1,B1)');
    expect(cell?.formula).toBe('CONCATENATE(A1,B1)');
    expect(cell?.computedValue).toBe('Hello World');
  });

  test('should parse formula cell with empty computed value', async () => {
    const xml = '<row><c><f>IF(A1&gt;0,A1,"")</f><v></v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=IF(A1>0,A1,"")');
    expect(cell?.formula).toBe('IF(A1>0,A1,"")');
    expect(cell?.computedValue).toBe(null);
  });

  test('should parse formula cell with zero computed value', async () => {
    const xml = '<row><c><f>SUM(A1:A1)</f><v>0</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=SUM(A1:A1)');
    expect(cell?.formula).toBe('SUM(A1:A1)');
    expect(cell?.computedValue).toBe(0);
  });

  test('should parse formula cell with decimal computed value', async () => {
    const xml = '<row><c><f>AVERAGE(A1:A3)</f><v>3.14159</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=AVERAGE(A1:A3)');
    expect(cell?.formula).toBe('AVERAGE(A1:A3)');
    expect(cell?.computedValue).toBe(3.14159);
  });

  test('should parse row with mixed formula and regular cells', async () => {
    const xml = '<row><c><v>10</v></c><c><f>SUM(A1:A1)</f><v>10</v></c><c><v>20</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe(10);
    expect(rows[0]?.cells[0]?.type).toBe('number');
    expect(rows[0]?.cells[1]?.type).toBe('formula');
    expect(rows[0]?.cells[1]?.value).toBe('=SUM(A1:A1)');
    expect(rows[0]?.cells[1]?.computedValue).toBe(10);
    expect(rows[0]?.cells[2]?.value).toBe(20);
    expect(rows[0]?.cells[2]?.type).toBe('number');
  });

  test('should parse formula cell without computed value element', async () => {
    // Formula cell without <v> element - can occur with manual calculation mode
    // or when formulas haven't been evaluated yet
    const xml = '<row><c><f>SUM(A1:A5)</f></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=SUM(A1:A5)');
    expect(cell?.formula).toBe('SUM(A1:A5)');
    expect(cell?.computedValue).toBeUndefined();
  });

  test('should parse formula cell with boolean computed value (TRUE)', async () => {
    // Formula that returns TRUE - stored as 1 in XLSX
    const xml = '<row><c t="b"><f>A1&gt;B1</f><v>1</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=A1>B1');
    expect(cell?.formula).toBe('A1>B1');
    expect(cell?.computedValue).toBe(true);
  });

  test('should parse formula cell with boolean computed value (FALSE)', async () => {
    // Formula that returns FALSE - stored as 0 in XLSX
    const xml = '<row><c t="b"><f>A1&gt;B1</f><v>0</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    const cell = rows[0]?.cells[0];
    expect(cell?.type).toBe('formula');
    expect(cell?.value).toBe('=A1>B1');
    expect(cell?.formula).toBe('A1>B1');
    expect(cell?.computedValue).toBe(false);
  });

  test('should parse row with attributes (row index)', async () => {
    const xml = '<row r="1"><c><v>Test</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows[0]?.rowIndex).toBe(1);
  });

  test('should parse multiple rows', async () => {
    const xml = '<worksheet><sheetData><row><c><v>Row1</v></c></row><row><c><v>Row2</v></c></row></sheetData></worksheet>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('Row1');
    expect(rows[1]?.cells[0]?.value).toBe('Row2');
  });

  test('should handle empty row', async () => {
    const xml = '<row></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { skipEmptyRows: false })) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(0);
  });

  test('should skip empty rows by default', async () => {
    // Create XML with empty rows between data rows
    const xmlWithEmptyRows = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Data1</t></is></c>
    </row>
    <row r="2">
      <!-- Empty row with no cells -->
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Data2</t></is></c>
    </row>
    <row r="4">
      <c r="A4"></c>
      <c r="B4"></c>
      <!-- Row with empty cells -->
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Data3</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithEmptyRows);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      parsedRows.push(row);
    }

    // Should skip empty rows (default behavior)
    expect(parsedRows).toHaveLength(3);
    expect(parsedRows[0]?.cells[0]?.value).toBe('Data1');
    expect(parsedRows[1]?.cells[0]?.value).toBe('Data2');
    expect(parsedRows[2]?.cells[0]?.value).toBe('Data3');
  });

  test('should include empty rows when skipEmptyRows is false', async () => {
    // Create XML with empty rows between data rows
    const xmlWithEmptyRows = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Data1</t></is></c>
    </row>
    <row r="2">
      <!-- Empty row with no cells -->
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Data2</t></is></c>
    </row>
    <row r="4">
      <c r="A4"></c>
      <c r="B4"></c>
      <!-- Row with empty cells -->
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Data3</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithEmptyRows);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { skipEmptyRows: false })) {
      parsedRows.push(row);
    }

    // Should include all rows, including empty ones
    expect(parsedRows).toHaveLength(5);
    expect(parsedRows[0]?.cells[0]?.value).toBe('Data1');
    // Row 2 is empty (no cells)
    expect(parsedRows[1]?.cells).toHaveLength(0);
    expect(parsedRows[2]?.cells[0]?.value).toBe('Data2');
    // Row 4 has empty cells
    expect(parsedRows[3]?.cells).toHaveLength(2);
    expect(parsedRows[3]?.cells[0]?.value).toBe('');
    expect(parsedRows[3]?.cells[1]?.value).toBe('');
    expect(parsedRows[4]?.cells[0]?.value).toBe('Data3');
  });

  test('should skip empty rows when reading XML (default behavior)', async () => {
    // Create XML with gaps in row numbers (simulating empty rows)
    const xmlWithEmptyRows = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Row1</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Row3</t></is></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Row5</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithEmptyRows);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      parsedRows.push(row);
    }

    // Should only return rows with data (default skipEmptyRows behavior)
    expect(parsedRows).toHaveLength(3);
    expect(parsedRows[0]?.cells[0]?.value).toBe('Row1');
    expect(parsedRows[1]?.cells[0]?.value).toBe('Row3');
    expect(parsedRows[2]?.cells[0]?.value).toBe('Row5');
  });

  test('should include empty rows when skipEmptyRows is false (with empty cells)', async () => {
    // Create XML with empty rows between data rows
    const xmlWithEmptyRows = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Row1</t></is></c>
    </row>
    <row r="2">
      <!-- Empty row -->
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Row3</t></is></c>
    </row>
    <row r="4">
      <c r="A4"></c>
      <!-- Row with empty cell -->
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>Row5</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithEmptyRows);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { skipEmptyRows: false })) {
      parsedRows.push(row);
    }

    // Should include all rows including empty ones
    expect(parsedRows).toHaveLength(5);
    expect(parsedRows[0]?.cells[0]?.value).toBe('Row1');
    expect(parsedRows[1]?.cells).toHaveLength(0); // Empty row
    expect(parsedRows[2]?.cells[0]?.value).toBe('Row3');
    expect(parsedRows[3]?.cells[0]?.value).toBe(''); // Row with empty cell
    expect(parsedRows[4]?.cells[0]?.value).toBe('Row5');
  });

  test('should not treat rows with null values as empty', async () => {
    // Test that rows with null values are NOT considered empty
    // Null represents invalid but present data (e.g., invalid dates), so the row should be returned
    const xmlWithNulls = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="d"><v>-700000</v></c>
      <!-- Invalid date becomes null -->
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Valid</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

    const bytes = async function* () {
      yield new TextEncoder().encode(xmlWithNulls);
    }();

    const parsedRows: any[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { use1904Dates: false })) {
      parsedRows.push(row);
    }

    // Row with null value should be returned (null represents invalid but present data)
    expect(parsedRows).toHaveLength(2);
    expect(parsedRows[0]?.cells[0]?.value).toBeNull();
    expect(parsedRows[1]?.cells[0]?.value).toBe('Valid');
  });

  test('should parse inline strings', async () => {
    const xml = '<row><c t="inlineStr"><is><t>Inline Text</t></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    expect(rows[0]?.cells[0]?.value).toBe('Inline Text');
    expect(rows[0]?.cells[0]?.type).toBe('string');
  });

  test('should parse mixed inline strings and regular values', async () => {
    const xml = '<row><c t="inlineStr"><is><t>Text</t></is></c><c><v>42</v></c><c t="inlineStr"><is><t>More</t></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('Text');
    expect(rows[0]?.cells[0]?.type).toBe('string');
    expect(rows[0]?.cells[1]?.value).toBe(42);
    expect(rows[0]?.cells[1]?.type).toBe('number');
    expect(rows[0]?.cells[2]?.value).toBe('More');
    expect(rows[0]?.cells[2]?.type).toBe('string');
  });

  test('should parse date cells with conversion', async () => {
    // Excel date 1 = January 1, 1900
    const xml = '<row><c t="d"><v>1</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes), undefined, { use1904Dates: false })) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBeInstanceOf(Date);
    expect((dateCell?.value as Date)?.getFullYear()).toBe(1900);
    expect((dateCell?.value as Date)?.getMonth()).toBe(0); // January
    expect((dateCell?.value as Date)?.getDate()).toBe(1);
  });

  test('should format dates when shouldFormatDates is true', async () => {
    // Excel date 42382 = January 13, 2016
    const xml = '<row><c t="d" s="1"><v>42382</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    // Create a style format map with format code MM/DD/YYYY
    const styleFormatMap = new Map<number, string>();
    styleFormatMap.set(1, 'MM/DD/YYYY');

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      styleFormatMap,
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBe('01/13/2016');
  });

  test('should return Date objects when shouldFormatDates is false', async () => {
    // Excel date 42382 = January 13, 2016
    const xml = '<row><c t="d" s="1"><v>42382</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const styleFormatMap = new Map<number, string>();
    styleFormatMap.set(1, 'MM/DD/YYYY');

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: false },
      styleFormatMap,
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBeInstanceOf(Date);
    expect((dateCell?.value as Date)?.getFullYear()).toBe(2016);
    expect((dateCell?.value as Date)?.getMonth()).toBe(0); // January
    expect((dateCell?.value as Date)?.getDate()).toBe(13);
  });

  test('should detect dates from numeric cells with date format codes', async () => {
    // Excel date 42382 = January 13, 2016
    // Cell has numeric value but date format code
    const xml = '<row><c s="1"><v>42382</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    // Create a style format map with date format code
    const styleFormatMap = new Map<number, string>();
    styleFormatMap.set(1, 'MM/DD/YYYY');

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      styleFormatMap,
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBe('01/13/2016');
  });

  test('should detect dates from numeric cells with date format codes', async () => {
    // Excel date 42382 = January 13, 2016
    // Cell has numeric value but date format code (no t="d" attribute)
    // This tests that date detection works regardless of shouldFormatDates
    const xml = '<row><c s="1"><v>42382</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    // Create a style format map with date format code
    const styleFormatMap = new Map<number, string>();
    styleFormatMap.set(1, 'MM/DD/YYYY');

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: false },
      styleFormatMap,
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    // Should return Date object, not formatted string
    expect(dateCell?.value).toBeInstanceOf(Date);
    expect((dateCell?.value as Date)?.getFullYear()).toBe(2016);
    expect((dateCell?.value as Date)?.getMonth()).toBe(0); // January
    expect((dateCell?.value as Date)?.getDate()).toBe(13);
  });

  test('should format dates with time format codes', async () => {
    // Excel date with time: 42382.2 = January 13, 2016 4:48:00 AM
    const xml = '<row><c s="1"><v>42382.2</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const styleFormatMap = new Map<number, string>();
    styleFormatMap.set(1, 'h:mm:ss AM/PM');

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      styleFormatMap,
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    // Should format as time (exact format may vary, but should contain time components)
    expect(typeof dateCell?.value).toBe('string');
    expect((dateCell?.value as string)).toMatch(/\d+:\d+:\d+/);
  });

  test('should use default format when format code is missing', async () => {
    // Excel date 42382 = January 13, 2016
    const xml = '<row><c t="d"><v>42382</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      new Map(),
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    // Should use default format yyyy-MM-dd
    expect(dateCell?.value).toBe('2016-01-13');
  });

  test('should handle ISO 8601 date strings when shouldFormatDates is false', async () => {
    const xml = '<row><c t="d"><v>1976-11-22T08:30:00.000</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: false },
      new Map(),
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    expect(dateCell?.value).toBeInstanceOf(Date);
    expect((dateCell?.value as Date)?.getFullYear()).toBe(1976);
    expect((dateCell?.value as Date)?.getMonth()).toBe(10); // November (0-indexed)
    expect((dateCell?.value as Date)?.getDate()).toBe(22);
  });

  test('should preserve ISO 8601 date strings when shouldFormatDates is true', async () => {
    const xml = '<row><c t="d"><v>1976-11-22T08:30:00.000</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      new Map(),
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    // Should preserve the original string
    expect(dateCell?.value).toBe('1976-11-22T08:30:00.000');
  });

  test('should handle invalid dates gracefully', async () => {
    // Invalid Excel date (out of range) - using value larger than max valid (2958465)
    const xml = '<row><c t="d"><v>3000000</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(
      parseXmlEvents(bytes),
      undefined,
      { use1904Dates: false, shouldFormatDates: true },
      new Map(),
    )) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    const dateCell = rows[0]?.cells[0];
    expect(dateCell?.type).toBe('date');
    // Invalid dates should return null
    expect(dateCell?.value).toBeNull();
  });

  test('should parse boolean cells', async () => {
    const xml = '<row><c t="b"><v>1</v></c><c t="b"><v>0</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(2);
    expect(rows[0]?.cells[0]?.type).toBe('boolean');
    expect(rows[0]?.cells[0]?.value).toBe(true);
    expect(rows[0]?.cells[1]?.type).toBe('boolean');
    expect(rows[0]?.cells[1]?.value).toBe(false);
  });

  test('should handle cells with explicit references out of order', async () => {
    // B1 comes first, then A1 - A1 should correctly overwrite the empty placeholder at index 0
    const xml = '<row><c r="B1"><v>B1Value</v></c><c r="A1"><v>A1Value</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(2);
    expect(rows[0]?.cells[0]?.value).toBe('A1Value'); // A1 at index 0
    expect(rows[0]?.cells[1]?.value).toBe('B1Value'); // B1 at index 1
  });

  test('should handle cells with explicit references in reverse order', async () => {
    // C1, B1, A1 in that order - all should be placed correctly
    const xml = '<row><c r="C1"><v>C1Value</v></c><c r="B1"><v>B1Value</v></c><c r="A1"><v>A1Value</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('A1Value'); // A1 at index 0
    expect(rows[0]?.cells[1]?.value).toBe('B1Value'); // B1 at index 1
    expect(rows[0]?.cells[2]?.value).toBe('C1Value'); // C1 at index 2
  });

  test('should handle mixed explicit and implicit cell references', async () => {
    // Mix of cells with and without explicit r attributes
    // B1 (explicit), then implicit cell, then A1 (explicit)
    const xml = '<row><c r="B1"><v>B1Value</v></c><c><v>ImplicitValue</v></c><c r="A1"><v>A1Value</v></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    // B1 creates [empty, B1Value], implicit appends [empty, B1Value, ImplicitValue], A1 overwrites index 0
    expect(rows[0]?.cells).toHaveLength(3);
    expect(rows[0]?.cells[0]?.value).toBe('A1Value'); // A1 at index 0 (overwrites empty placeholder)
    expect(rows[0]?.cells[1]?.value).toBe('B1Value'); // B1 at index 1
    expect(rows[0]?.cells[2]?.value).toBe('ImplicitValue'); // Implicit cell at index 2
  });

  test('should accumulate text from multiple rich text runs in inline strings', async () => {
    // Inline string with multiple rich text runs (<r> elements)
    // Each <r><t>...</t></r> should contribute to the final value
    // Note: Excel stores exactly what's typed - if there's no space in the XML, there's no space in the result
    const xml = '<row><c t="inlineStr"><is><r><t>Bold</t></r><r><t>Normal</t></r></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    // Should accumulate "Bold" + "Normal" = "BoldNormal" (no space - Excel stores exactly what's in XML)
    expect(rows[0]?.cells[0]?.value).toBe('BoldNormal');
    expect(rows[0]?.cells[0]?.type).toBe('string');
  });

  test('should preserve spaces in rich text runs when present in XML', async () => {
    // Inline string with multiple rich text runs where spaces are part of the text content
    // This demonstrates that Excel stores exactly what's typed - spaces must be in the XML
    const xml = '<row><c t="inlineStr"><is><r><t>Bold </t></r><r><t>Normal</t></r></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    // Should accumulate "Bold " + "Normal" = "Bold Normal" (space preserved from XML)
    expect(rows[0]?.cells[0]?.value).toBe('Bold Normal');
    expect(rows[0]?.cells[0]?.type).toBe('string');
  });

  test('should handle inline string with single text element (no rich text runs)', async () => {
    // Inline string with <t> directly under <is> (no <r> wrapper)
    const xml = '<row><c t="inlineStr"><is><t>Simple Text</t></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    expect(rows[0]?.cells[0]?.value).toBe('Simple Text');
    expect(rows[0]?.cells[0]?.type).toBe('string');
  });

  test('should handle inline string with multiple rich text runs and empty text', async () => {
    // Test accumulation with empty text elements
    const xml = '<row><c t="inlineStr"><is><r><t>First</t></r><r><t></t></r><r><t>Last</t></r></is></c></row>';
    const bytes = async function* () {
      yield new TextEncoder().encode(xml);
    }();

    const rows: Row[] = [];
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(1);
    // Should accumulate "First" + "" + "Last" = "FirstLast"
    expect(rows[0]?.cells[0]?.value).toBe('FirstLast');
    expect(rows[0]?.cells[0]?.type).toBe('string');
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
});
