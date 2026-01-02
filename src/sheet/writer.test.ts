import { describe, test, expect } from 'bun:test';
import { writeSheetXml } from '@xml/writer';
import type { Row } from '../types';

describe('Row Writer Integration', () => {
  test('should write single row', async () => {
    const rows: Row[] = [{ cells: [{ value: 'Hello', type: 'string' }], rowIndex: 1 }];
    const chunks: string[] = [];
    for await (const chunk of writeSheetXml(async function* () {
      for (const row of rows) yield row;
    }())) {
      chunks.push(chunk);
    }
    const result = chunks.join('');
    expect(result).toContain('<row r="1" spans="1:1">');
    expect(result).toContain('<c r="A1" t="inlineStr"><is><t>Hello</t></is></c>');
    expect(result).toContain('</row>');
  });

  test('should write multiple rows', async () => {
    const rows: Row[] = [
      { cells: [{ value: 'A', type: 'string' }] },
      { cells: [{ value: 'B', type: 'string' }] },
      { cells: [{ value: 'C', type: 'string' }] },
    ];
    const chunks: string[] = [];
    for await (const chunk of writeSheetXml(async function* () {
      for (const row of rows) yield row;
    }())) {
      chunks.push(chunk);
    }
    const result = chunks.join('');
    const rowMatches = result.match(/<row r="/g);
    expect(rowMatches).toHaveLength(3);
  });

  test('should write rows with different cell types', async () => {
    const rows: Row[] = [
      {
        cells: [
          { value: 'Name', type: 'string' },
          { value: 42, type: 'number' },
          { value: 'Text', type: 'string' },
        ],
        rowIndex: 1,
      },
    ];
    const chunks: string[] = [];
    for await (const chunk of writeSheetXml(async function* () {
      for (const row of rows) yield row;
    }())) {
      chunks.push(chunk);
    }
    const result = chunks.join('');
    expect(result).toContain('<c r="A1" t="inlineStr"><is><t>Name</t></is></c>');
    expect(result).toContain('<c r="B1"><v>42</v></c>');
    expect(result).toContain('<c r="C1" t="inlineStr"><is><t>Text</t></is></c>');
  });

  test('should verify XML output format matches Excel spec', async () => {
    const rows: Row[] = [
      {
        cells: [{ value: 'Test', type: 'string' }],
        rowIndex: 1,
      },
    ];
    const chunks: string[] = [];
    for await (const chunk of writeSheetXml(async function* () {
      for (const row of rows) yield row;
    }())) {
      chunks.push(chunk);
    }
    const result = chunks.join('');
    // Should have proper XML structure
    expect(result).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(result).toContain('<worksheet');
    expect(result).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(result).toMatch(/<sheetData>/);
    expect(result).toMatch(/<row r="1" spans="1:1">/);
    expect(result).toMatch(/<\/sheetData>/);
    expect(result).toMatch(/<\/worksheet>/);
  });

  test('should handle empty rows', async () => {
    const rows: Row[] = [{ cells: [], rowIndex: 1 }];
    const chunks: string[] = [];
    for await (const chunk of writeSheetXml(async function* () {
      for (const row of rows) yield row;
    }())) {
      chunks.push(chunk);
    }
    const result = chunks.join('');
    expect(result).toContain('<row r="1" spans="1:1"></row>');
  });
});
