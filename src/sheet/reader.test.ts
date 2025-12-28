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
    for await (const row of parseSheet(parseXmlEvents(bytes))) {
      rows.push(row);
    }

    expect(rows).toHaveLength(1);
    expect(rows[0]?.cells).toHaveLength(0);
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
});

