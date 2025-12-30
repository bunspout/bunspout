
import { describe, test, expect } from 'bun:test';
import { writeSheetXml, serializeRow, serializeCell } from './writer';
import type { Row, Cell } from '../types';

describe('XML Writer', () => {
  describe('serializeCell', () => {
    test('should serialize string cell', () => {
      const cell: Cell = { value: 'Hello', type: 'string' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0" t="inlineStr"><is><t>Hello</t></is></c>');
    });

    test('should serialize number cell', () => {
      const cell: Cell = { value: 42, type: 'number' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0"><v>42</v></c>');
    });

    test('should serialize cell without explicit type (defaults to number)', () => {
      const cell: Cell = { value: 100 };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0"><v>100</v></c>');
    });

    test('should escape XML special characters in string values', () => {
      const cell: Cell = { value: '<test>&"value"</test>', type: 'string' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toContain('&lt;test&gt;&amp;&quot;value&quot;&lt;/test&gt;');
    });

    test('should serialize date cell', () => {
      const cell: Cell = { value: 45323, type: 'date' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0" t="d"><v>45323</v></c>');
    });

    test('should serialize boolean cell (true)', () => {
      const cell: Cell = { value: 1, type: 'boolean' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0" t="b"><v>1</v></c>');
    });

    test('should serialize boolean cell (false)', () => {
      const cell: Cell = { value: 0, type: 'boolean' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toBe('<c r="A1" s="0" t="b"><v>0</v></c>');
    });

    test('should preserve newlines in string values', () => {
      const cell: Cell = { value: 'Line1\nLine2\r\nLine3', type: 'string' };
      const result = serializeCell(cell, 1, 0);
      // Newlines should be preserved (not escaped)
      expect(result).toContain('Line1\nLine2\r\nLine3');
      expect(result).toBe('<c r="A1" s="0" t="inlineStr"><is><t>Line1\nLine2\r\nLine3</t></is></c>');
    });

    test('should preserve tabs in string values', () => {
      const cell: Cell = { value: 'Col1\tCol2', type: 'string' };
      const result = serializeCell(cell, 1, 0);
      expect(result).toContain('Col1\tCol2');
      expect(result).toBe('<c r="A1" s="0" t="inlineStr"><is><t>Col1\tCol2</t></is></c>');
    });

    test('should handle control characters by removing invalid ones', () => {
      // Control characters: null (0x00), bell (0x07), form feed (0x0C), etc.
      // Valid: tab (0x09), newline (0x0A), carriage return (0x0D)
      const cell: Cell = {
        value: 'Text' + String.fromCharCode(0x00) + 'Middle' + String.fromCharCode(0x07) + 'End',
        type: 'string',
      };
      const result = serializeCell(cell, 1, 0);
      // Invalid control chars should be removed, but text preserved
      expect(result).toContain('Text');
      expect(result).toContain('Middle');
      expect(result).toContain('End');
      // Should not contain the null or bell characters
      expect(result).not.toContain(String.fromCharCode(0x00));
      expect(result).not.toContain(String.fromCharCode(0x07));
    });

    test('should generate correct cell references', () => {
      const cell: Cell = { value: 'Test', type: 'string' };
      expect(serializeCell(cell, 1, 0)).toContain('r="A1"');
      expect(serializeCell(cell, 1, 1)).toContain('r="B1"');
      expect(serializeCell(cell, 2, 0)).toContain('r="A2"');
      expect(serializeCell(cell, 26, 0)).toContain('r="A26"');
      expect(serializeCell(cell, 27, 0)).toContain('r="A27"');
    });
  });

  describe('serializeRow', () => {
    test('should serialize row with one cell', () => {
      const row: Row = { cells: [{ value: 'Hello', type: 'string' }], rowIndex: 1 };
      const result = serializeRow(row);
      expect(result).toContain('<row r="1" spans="1:1">');
      expect(result).toContain('</row>');
      expect(result).toContain('<c r="A1" s="0" t="inlineStr"><is><t>Hello</t></is></c>');
    });

    test('should serialize row with multiple cells', () => {
      const row: Row = {
        cells: [
          { value: 'A', type: 'string' },
          { value: 1, type: 'number' },
          { value: 'B', type: 'string' },
        ],
        rowIndex: 1,
      };
      const result = serializeRow(row);
      expect(result).toContain('<row r="1" spans="1:3">');
      expect(result).toContain('<c r="A1" s="0" t="inlineStr"><is><t>A</t></is></c>');
      expect(result).toContain('<c r="B1" s="0"><v>1</v></c>');
      expect(result).toContain('<c r="C1" s="0" t="inlineStr"><is><t>B</t></is></c>');
    });

    test('should serialize row with rowIndex attribute', () => {
      const row: Row = { cells: [{ value: 'Test' }], rowIndex: 5 };
      const result = serializeRow(row);
      expect(result).toContain('<row r="5" spans="1:1">');
    });

    test('should calculate spans correctly for row with holes', () => {
      // Row with cells at indices 2 and 5, but not at 0, 1, 3, 4
      const row: Row = {
        cells: [
          undefined,
          undefined,
          { value: 'C' },
          undefined,
          undefined,
          { value: 'F' },
        ],
      };
      const result = serializeRow(row);
      // Should span from column 3 (index 2 + 1) to column 6 (index 5 + 1)
      expect(result).toContain('spans="3:6"');
      expect(result).toContain('<c r="C1"');
      expect(result).toContain('<c r="F1"');
    });

    test('should handle empty row (all undefined)', () => {
      const row: Row = {
        cells: [undefined, undefined, undefined],
      };
      const result = serializeRow(row);
      // Should default to spans="1:1" when no cells found
      expect(result).toContain('spans="1:1"');
    });
  });

  describe('writeSheetXml', () => {
    test('should write empty sheet', async () => {
      const rows: Row[] = [];
      const chunks: string[] = [];
      for await (const chunk of writeSheetXml(async function* () {
        for (const row of rows) yield row;
      }())) {
        chunks.push(chunk);
      }
      const result = chunks.join('');
      expect(result).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
      expect(result).toContain('<worksheet');
      expect(result).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
      expect(result).toContain('<sheetData>');
      expect(result).toContain('</sheetData>');
      expect(result).toContain('</worksheet>');
    });

    test('should write sheet with one row', async () => {
      const rows: Row[] = [{ cells: [{ value: 'Hello', type: 'string' }], rowIndex: 1 }];
      const chunks: string[] = [];
      for await (const chunk of writeSheetXml(async function* () {
        for (const row of rows) yield row;
      }())) {
        chunks.push(chunk);
      }
      const result = chunks.join('');
      expect(result).toContain('<row r="1" spans="1:1">');
      expect(result).toContain('<c r="A1" s="0" t="inlineStr"><is><t>Hello</t></is></c>');
      expect(result).toContain('</row>');
    });

    test('should write sheet with multiple rows', async () => {
      const rows: Row[] = [
        { cells: [{ value: 'A', type: 'string' }], rowIndex: 1 },
        { cells: [{ value: 'B', type: 'string' }], rowIndex: 2 },
        { cells: [{ value: 'C', type: 'string' }], rowIndex: 3 },
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

    test('should stream rows incrementally', async () => {
      let rowCount = 0;
      const rows: Row[] = [
        { cells: [{ value: 'A' }] },
        { cells: [{ value: 'B' }] },
      ];
      for await (const chunk of writeSheetXml(async function* () {
        for (const row of rows) {
          rowCount++;
          yield row;
        }
      }())) {
        // Just consume chunks
      }
      expect(rowCount).toBe(2);
    });
  });
});
