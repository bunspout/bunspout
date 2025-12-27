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
});

